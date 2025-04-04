using System.Reflection;
using System.Text;

namespace ObjectToCsv
{

    public static class Conversion
    {
        private static string _dateFormat  = "";
        /// <summary>
        /// Converts a list of objects to a CSV-formatted string.
        /// </summary>
        /// <typeparam name="T">The type of objects in the list.</typeparam>
        /// <param name="objects">The list of objects to convert.</param>
        /// <param name="dateFormat">An optional custom date/time format string. Defaults to "dd MMMM yyyy".</param>
        /// <returns>The CSV-formatted string representing the object list.</returns>
        public static string BindCsv<T>(List<T> objects,string dateFormat="")
        {
            _dateFormat = string.IsNullOrEmpty(dateFormat) ? "dd MMMM yyyy" : dateFormat;
            string csvStr = string.Empty;
            if (objects.Count == 0) return string.Empty;

            var modelType = objects?.FirstOrDefault().GetType();
            var properties = modelType.GetProperties();
            properties.ToList().ForEach(p =>
            {
                var propAttribute = p.GetCustomAttributes(typeof(Header), false).FirstOrDefault() as Header;
                var ordinal = p.GetCustomAttributes(typeof(Ordinal), false).FirstOrDefault() as Ordinal;
                if (propAttribute == null || propAttribute.HeaderName == string.Empty)//check if Header available
                {
                    throw new Exception("Please add Header attribute");
                }
                if (ordinal == null)//check if Ordinal available
                {
                    throw new Exception("Please add Header attribute");
                }
                csvStr += $"\"{propAttribute.HeaderName}\",";

            });
            csvStr = csvStr.Remove(csvStr.Length - 1);
            csvStr += Environment.NewLine;
            objects.ForEach(o =>
            {
                csvStr += BindRow<T>(o);
            });
            return csvStr;
        }
        public static MemoryStream GetCsvMemoryStream<T>(List<T> objects)
        {
            string csvStr = BindCsv<T>(objects);
            MemoryStream ms = new MemoryStream();
            using (StreamWriter sw = new StreamWriter(ms, new UTF8Encoding(false), -1, true))
            {
                sw.WriteLine(csvStr);
            }
            ms.Position = 0;
            return ms;
        }
        private static string BindRow<T>(T model)
        {
            string csvRowStr = string.Empty;
            var modelType = model.GetType();
            var properties = modelType.GetProperties();
            string[] rowArr = new string[properties.Length];
            properties.ToList().ForEach(p =>
            {
                var ordinal = p.GetCustomAttributes(typeof(Ordinal), false).FirstOrDefault() as Ordinal;
                object val = p?.GetValue(model, null); ;
                if ((p.PropertyType == typeof(DateTime) || p.PropertyType == typeof(DateTime?)) && val != null)
                {
                    val = ((DateTime)p?.GetValue(model, null)).ToString(_dateFormat);
                }
                else if (p.PropertyType == typeof(string) && val == null)
                {
                    val = string.Empty;
                }
                rowArr[ordinal.Order] = val==null ? "": val.ToString();

            });
            rowArr.ToList().ForEach(row =>
            {
                string temp = row.Contains(',') ? $"\"{row}\"" : row;
                csvRowStr += $"{temp},";
            });
            csvRowStr = csvRowStr.Remove(csvRowStr.Length - 1);
            csvRowStr += Environment.NewLine;
            return csvRowStr;
        }

    }

}
