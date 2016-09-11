using System;

namespace DataFetcherAddIn
{
    [Serializable]
    public class DataHolder
    {
        public DateTime Date { get; set; }
        public string Alias1 { get; set; }
        public string Alias2 { get; set; }
        public string Alias3 { get; set; }
        public string Alias4 { get; set; }
        public int InsertMode { get; set;}

        public static DataHolder Default = new DataHolder
        {
            Date = DateTime.Now.Date,
            InsertMode = 0
        };
    }
}
