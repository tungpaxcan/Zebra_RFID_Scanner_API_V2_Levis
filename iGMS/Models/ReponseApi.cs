using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Web;

namespace Zebra_RFID_Scanner.Models
{

    public class ReponseApi
    {

        public int statusCode { get; set; }

        public string body { get; set; }

        public bool isBase64Encoded { get; set; }
    }

    public class ReponseCheckFile
    {
        public float code { get; set; }
        public string url { get; set; }
        public bool status { get; set; }
    }
    public class Certificate2
    {
        public string msg { get; set; }

        public X509Certificate2 cert { get; set; }
    }
}