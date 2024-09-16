using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Zebra_RFID_Scanner.Models;

namespace Zebra_RFID_Scanner.Controllers
{
    public class HODController : BaseController
    {
        private Entities db = new Entities();
        // GET: HOD
        public ActionResult Index()
        {
            return View();
        }
        public ActionResult Levis2()
        {
            return View();
        }
        public ActionResult Uniqlo()
        {
            return View();
        }
        [HttpPost]
        public JsonResult DeleteCtn(int id,string idName)
        {
            try
            {
            
                var ctn = db.DataScanPhysicals.Find(id);
                db.DataScanPhysicals.Remove(ctn);
                db.SaveChanges();
                var Ctn = (from a in db.DataScanPhysicals.Where(x => x.IdReports == idName)
                           select new
                           {
                               so = a.So,
                               po = a.Po,
                               sku = a.Sku,
                               id = a.Id,
                               ctn = a.CartonTo,
                               status = a.Status == true ? "Matched" : "Mismatched",
                               upc = a.UPC,
                               qty = a.Qty
                           }).ToList();
                return Json(new { code = 200, Ctn = Ctn ,}, JsonRequestBehavior.AllowGet);
            }
            catch (Exception e)
            {
                return Json(new { code = 500, msg ="false" + e.Message }, JsonRequestBehavior.AllowGet);
            }
        }
        [HttpPost]
        public JsonResult FixCtn(int id)
        {
            try
            {
                var data = db.Discrepancies.Find(id);
                var idReports = data.IdReports;
                var listMatch = db.Discrepancies.Where(x => x.IdReports == idReports).ToList();
                listMatch.Remove(data);
                var totalQty = listMatch.Sum(x => int.Parse(x.Qty));
                data.QtyScan = (int.Parse(data.Qty) - (int.Parse(totalQty.ToString()) - int.Parse(data.QtyScan))).ToString();
                var generalList = listMatch.Select(x => new General { 
                   IdReports =  x.IdReports,
                   So = x.So,
                   Po = x.Po,
                   Sku = x.Sku,
                   CartonTo = x.CartonTo,
                   UPC = x.UPC,
                   Qty = x.Qty,
                   CreateDate = x.CreateDate,
                   CreateBy = x.CreateBy,
                   ModifyDate  =x.ModifyDate,
                   ModifyBy = x.ModifyBy,
                   Status = true
                }).ToList();
                db.Generals.AddRange(generalList);
                db.Discrepancies.RemoveRange(listMatch);
                db.SaveChanges();
                return Json(new { code = 200 }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception e)
            {
                return Json(new { code = 500, msg = "false" + e.Message }, JsonRequestBehavior.AllowGet);
            }
        }
    }
}