using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Zebra_RFID_Scanner.Controllers
{
    public class DataApi
    {
        public string PO { get; set; }
        public string Ctn { get; set; }
        public string EPC { get; set; }
        public string UPC { get; set; }
    }
}