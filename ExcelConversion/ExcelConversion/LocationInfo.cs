namespace ExcelConversion
{
    public class LocationInfo
    {
        public string property { get; set; }
        public string bedrooms { get; set; }
        public string bathrooms { get; set; }
        public string dateOnMkt { get; set; }
        public string address { get; set; }
        public string cityState { get; set; }
    }

    public class RowColIndexes
    {
        public int RowIndex { get; set; }
        public int ColIndex { get; set; }
    }

    //public class MapVal {
    //    public string fieldName { get; set; }
    //    public string fieldLabel { get; set; }

    //    public int relativePos { get; set; }

    //    public int offset { get; set; }
    //}

    //public enum RelativePos
    //{
    //   D = 0,
    //   R = 1,
    //   U = 2,
    //   L = 3
    //}
}