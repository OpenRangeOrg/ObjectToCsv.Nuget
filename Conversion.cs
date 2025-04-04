using System.Reflection;
using System.Text;

namespace ObjectToCsv
{

    public static class Conversion
    {
        private static string _dateFormat  = "";
        private static string _delimiter = "";
        /// <summary>
        /// Converts a list of objects to a CSV-formatted string.
        /// </summary>
        /// <typeparam name="T">The type of objects in the list.</typeparam>
        /// <param name="objects">The list of objects to convert.</param>
        /// <param name="dateFormat">An optional custom date/time format string. Defaults to "dd MMMM yyyy".</param>
        /// <param name="delimiter">An optional custom csv delimiter. Defaults to ",".</param>
        /// <returns>The CSV-formatted string representing the object list.</returns>
        public static string BindCsv<T>(List<T> objects,string dateFormat="",string delimiter=",")
        {
            _dateFormat = string.IsNullOrEmpty(dateFormat) ? "dd MMMM yyyy" : dateFormat;
            _delimiter = delimiter;
            string csvStr = string.Empty;
            if (objects.Count == 0) return string.Empty;

            PropertyInfo[] properties = GetSortedProperties<T>("Ordinal","Order");
            properties.ToList().ForEach(p =>
            {
                
                var matchingAttribute = GetAttributesBySubstring(p, "Header").First();
                if (matchingAttribute != null)
                {
                    var nameProperty = matchingAttribute.GetType().GetProperty("HeaderName");
                    if (nameProperty != null)
                    {
                        var headerName = nameProperty.GetValue(matchingAttribute)?.ToString();
                        csvStr += $"\"{headerName}\"{_delimiter}";
                    }

                }
                

            });
            csvStr = csvStr.Remove(csvStr.Length - 1);
            csvStr += Environment.NewLine;
            objects.ForEach(o =>
            {
                csvStr += BindRow<T>(o);
            });
            return csvStr;
        }

        /// <summary>
        /// Convert a list of objects into a CSV file stored in a MemoryStream
        /// </summary>
        /// <typeparam name="T">The type of objects in the list.</typeparam>
        /// <param name="objects">The list of objects to convert.</param>
        /// <param name="dateFormat">An optional custom date/time format string. Defaults to "dd MMMM yyyy".</param>
        /// <param name="delimiter">An optional custom csv delimiter. Defaults to ",".</param>
        /// <returns>The CSV file MemoryStream representing the object list.</returns>
        public static MemoryStream GetCsvMemoryStream<T>(List<T> objects, string dateFormat = "", string delimiter = ",", Encoding encoding=null)
        {
            encoding = encoding ?? Encoding.UTF8;
            string csvStr = BindCsv<T>(objects, dateFormat, delimiter);
            MemoryStream ms = new MemoryStream();
            using (StreamWriter sw = new StreamWriter(ms, encoding, -1, true))
            {
                sw.WriteLine(csvStr);
            }
            ms.Position = 0;
            return ms;
        }
        /// <summary>
        /// Convert a list of objects into a CSV file stored in a TextWriter
        /// </summary>
        /// <typeparam name="T">The type of objects in the list.</typeparam>
        /// <param name="objects">The list of objects to convert.</param>
        /// <param name="dateFormat">An optional custom date/time format string. Defaults to "dd MMMM yyyy".</param>
        /// <param name="delimiter">An optional custom csv delimiter. Defaults to ",".</param>
        /// <returns>The CSV file TextWriter representing the object list.</returns>
        public static TextWriter GetCsvTextWriter<T>(List<T> objects, string dateFormat = "", string delimiter = ",")
        {
            string csvStr = BindCsv<T>(objects, dateFormat, delimiter);
            var stringWriter = new StringWriter();
            stringWriter.Write(csvStr);
            return stringWriter;
        }
        private static string BindRow<T>(T model)
        {
            string csvRowStr = string.Empty;
            var modelType = model.GetType();
            var properties = modelType.GetProperties();
            string[] rowArr = new string[properties.Length];
            List<ValueOrder> valueOrders = new List<ValueOrder>();
            properties.ToList().ForEach(p =>
            {
                int order =0;
                var matchingAttribute = GetAttributesBySubstring(p, "Ordinal").First();
                if (matchingAttribute != null)
                {
                    var nameProperty = matchingAttribute.GetType().GetProperty("Order");
                    if (nameProperty != null)
                    {
                        order = int.Parse(nameProperty.GetValue(matchingAttribute)?.ToString());
                    }
                }
                object val = p?.GetValue(model, null); ;
                if ((p.PropertyType == typeof(DateTime) || p.PropertyType == typeof(DateTime?)) && val != null)
                {
                    val = ((DateTime)p?.GetValue(model, null)).ToString(_dateFormat);
                }
                else if (p.PropertyType == typeof(string) && val == null)
                {
                    val = string.Empty;
                }
                string finalValue = val==null ? "": val.ToString();

                valueOrders.Add(new ValueOrder { Value = finalValue, Order = order });

            });
            valueOrders.OrderBy(v=> v.Order).ToList().ForEach(row =>
            {
                string temp = row.Value.Contains($"{_delimiter}") ? $"\"{row.Value}\"" : row.Value;
                csvRowStr += $"{temp}{_delimiter}";
            });
            csvRowStr = csvRowStr.Remove(csvRowStr.Length - 1);
            csvRowStr += Environment.NewLine;
            return csvRowStr;
        }
        public static IEnumerable<object> GetAttributesBySubstring(PropertyInfo property, string attributeSubstring)
        {
            return property.GetCustomAttributes(inherit: true)
                           .Where(attr => attr.GetType().Name.Contains(attributeSubstring, StringComparison.OrdinalIgnoreCase));
        }
        public static PropertyInfo[] GetSortedProperties<T>(string attributeName,string attributeVariableName)
        {
            // Step 1: Get all properties of type T
            PropertyInfo[] allProperties = typeof(T).GetProperties();

            // Step 3: Create a list to store properties with their Order values
            var propertyList = new List<(PropertyInfo Property, int? Order)>();

            foreach (var property in allProperties)
            {
                // Step 4: Retrieve the custom attribute
                var attribute = GetAttributesBySubstring(property, attributeName).First();
                if (attribute != null)
                {
                    // Step 5: Dynamically access the 'Order' property of the attribute
                    var orderProperty = attribute.GetType().GetProperty(attributeVariableName);
                    if (orderProperty != null)
                    {
                        var orderValue = (int?)orderProperty.GetValue(attribute);
                        propertyList.Add((property, orderValue));
                    }
                }
                else
                {
                    propertyList.Add((property, 0));
                }
            }

            return propertyList.OrderBy(p => p.Order).Select(x => x.Property).ToArray();

        }
    }
    internal class ValueOrder
    {
        public int Order { get; set; }
        public string Value { get; set; }
    }
}
