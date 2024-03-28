using Microsoft.Ajax.Utilities;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.NetworkInformation;
using System.Resources;
using System.Security.Cryptography.X509Certificates;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using Zebra_RFID_Scanner.Models;

namespace Zebra_RFID_Scanner.Controllers
{
    public class LoginController : Controller
    {
        
        private Entities db = new Entities();
        ResourceManager rm = new ResourceManager("Zebra_RFID_Scanner.Resources.Resource", typeof(Resources.Resource).Assembly);
        // GET: Login
        public ActionResult Index()
        {
            if (Session["user"] != null)
            {
                return RedirectToAction("Login", "Login");
            }
            return View();
        }
        public ActionResult Login()
        {
          
            return View();
        }
     

        [HttpGet]
        public async Task<JsonResult> LoginiGMS(string user,string pass)
        {
            try
            {
                //NetworkInterface[] nics = NetworkInterface.GetAllNetworkInterfaces();
                //string MAC ="";
                //for(var i =0; i < nics.Length;i++)
                //{
                //    if (nics[i].GetPhysicalAddress().ToString().Length != 0)
                //    {
                //        MAC = nics[i].GetPhysicalAddress().ToString();
                //    }
                //}

                //if (Encode.ToMD5(MAC) == Encode.Mac)
                //{

                var passEncode = Encode.ToMD5(pass);
                var a = db.Users.FirstOrDefault(x => x.User1 == user && x.Pass == passEncode);
                if (a != null)
                {

                    Session["user"] = a;
                    return Json(new { code = 200, Url = "/123-rfid-scanner/Index" }, JsonRequestBehavior.AllowGet);
                }
                else
                {
                    return Json(new { code = 300, msg = "Incorrect Account or Password !!!" }, JsonRequestBehavior.AllowGet);
                }

                //}
                //return Json(new { code = 400, msg = "Can not login" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception e)
            {
                return Json(new { code = 500, msg = rm.GetString("false").ToString() + e.InnerException.Message }, JsonRequestBehavior.AllowGet);
            }
        }
        [HttpGet]
        public JsonResult SignOut()
        {
            try
            {
                    Session["user"] = null;
                    var language =  Session["CurrentCulture"];
                return Json(new { code = 200, Url = "/Login/Login" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception e)
            {
                return Json(new { code = 500, msg = rm.GetString("false").ToString()+ " "+ e.Message }, JsonRequestBehavior.AllowGet);
            }
        }
       
    }
}