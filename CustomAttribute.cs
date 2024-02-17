namespace ObjectToCsv
{
    public class Header : Attribute
    {
        public string HeaderName { get; private set; }
        public Header(string headerName)
        {
            HeaderName = headerName;
        }
    }
    public class Ordinal : Attribute
    {
        public int Order { get; private set; }
        public Ordinal(int order)
        {
            Order = order;
        }
    }
}
