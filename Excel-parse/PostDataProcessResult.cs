using System;
using System.Collections.Generic;
using System.Text;

namespace Excel_parse
{
    public class PostDataProcessResultSet
    {
        public List<PostDataProcessResult> PostDataProcessResultList { get;set;}
    }
    public class PostDataProcessResult
    {
        public int RowNumber { get; set; }
        public string ItemId { get; set; }
        public string QRCodeImage { get; set; }
        public bool IsProcessed { get; set; }
    }
}
