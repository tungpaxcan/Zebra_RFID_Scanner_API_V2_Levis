using System;
using System.Collections.Generic;
using System.Linq;
using System.Resources;
using System.Web;
using System.Web.Mvc;
using Zebra_RFID_Scanner.Models;

namespace Zebra_RFID_Scanner.Controllers
{
    public class ReportsUNIQLOController : BaseController
    {
        private Entities db = new Entities();
        ResourceManager rm = new ResourceManager("Zebra_RFID_Scanner.Resources.Resource", typeof(Resources.Resource).Assembly);
        // GET: ReportsUNIQLO
        public ActionResult Index()
        {
            return View();
        }
        [HttpGet]
        public JsonResult ReportsAll(string name, string date)
        {
            try
            {
                string date1 = date == "" ? "" : date.Substring(0, date.IndexOf("/")).Trim();
                string date2 = date == "" ? "" : date.Substring(date.IndexOf("/") + 1).Trim();
                var s1 = sumDate(date1);
                var s2 = sumDate(date2);
                var reports = (from b in db.Reports
                               where b.Status == true && b.So == ""
                               select new
                               {
                                   id = b.Id,
                                   createDate = b.CreateDate,
                                   createBy = b.CreateBy,
                                   status = b.Status == true ? "Confirmed" : "Unconfimred"
                               }).ToList();
                if (name != "" && date != "")
                {
                    if (name == "-1")
                    {
                        var reportss = reports.Where(x => (Convert.ToInt32(x.createDate.Value.Day) + (Convert.ToInt32(x.createDate.Value.Month) * 30) + Convert.ToInt32(x.createDate.Value.Year)) >= s1 &&
                                                         (Convert.ToInt32(x.createDate.Value.Day) + (Convert.ToInt32(x.createDate.Value.Month) * 30) + Convert.ToInt32(x.createDate.Value.Year)) <= s2);
                        return Json(new { code = 200, reports = reportss }, JsonRequestBehavior.AllowGet);
                    }
                    else
                    {
                        var user = db.Users.SingleOrDefault(x => x.Id.ToString() == name);
                        name = user.Name;
                        var reportss = reports.Where(x => (Convert.ToInt32(x.createDate.Value.Day) + (Convert.ToInt32(x.createDate.Value.Month) * 30) + Convert.ToInt32(x.createDate.Value.Year)) >= s1 && x.createBy == name
                        && (Convert.ToInt32(x.createDate.Value.Day) + (Convert.ToInt32(x.createDate.Value.Month) * 30) + Convert.ToInt32(x.createDate.Value.Year)) <= s2);
                        return Json(new { code = 200, reports = reportss }, JsonRequestBehavior.AllowGet);
                    }
                }

                return Json(new { code = 200, reports = reports }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception e)
            {
                return Json(new { code = 500, msg = rm.GetString("false") + " " + e.Message }, JsonRequestBehavior.AllowGet);
            }
        }
        [HttpGet]
        public JsonResult UnconfirmedReports(string name, string date)
        {
            try
            {
                string date1 = date == "" ? "" : date.Substring(0, date.IndexOf("/")).Trim();
                string date2 = date == "" ? "" : date.Substring(date.IndexOf("/") + 1).Trim();
                var s1 = sumDate(date1);
                var s2 = sumDate(date2);
                var reports = (from b in db.Reports.Where(x => x.Status == false &&x.So == "")
                               select new
                               {
                                   id = b.Id,
                                   createDate = b.CreateDate,
                                   createBy = b.CreateBy,

                               }).ToList();
                if (name != "" && date != "")
                {
                    if (name == "-1")
                    {
                        var reportss = reports.Where(x => (Convert.ToInt32(x.createDate.Value.Day) + (Convert.ToInt32(x.createDate.Value.Month) * 30) + Convert.ToInt32(x.createDate.Value.Year)) >= s1 &&
                                                         (Convert.ToInt32(x.createDate.Value.Day) + (Convert.ToInt32(x.createDate.Value.Month) * 30) + Convert.ToInt32(x.createDate.Value.Year)) <= s2);
                        return Json(new { code = 200, reports = reportss }, JsonRequestBehavior.AllowGet);
                    }
                    else
                    {
                        var user = db.Users.SingleOrDefault(x => x.Id.ToString() == name);
                        name = user.Name;
                        var reportss = reports.Where(x => (Convert.ToInt32(x.createDate.Value.Day) + (Convert.ToInt32(x.createDate.Value.Month) * 30) + Convert.ToInt32(x.createDate.Value.Year)) >= s1 && x.createBy == name
                        && (Convert.ToInt32(x.createDate.Value.Day) + (Convert.ToInt32(x.createDate.Value.Month) * 30) + Convert.ToInt32(x.createDate.Value.Year)) <= s2);
                        return Json(new { code = 200, reports = reportss }, JsonRequestBehavior.AllowGet);
                    }
                }
                return Json(new { code = 200, reports = reports }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception e)
            {
                return Json(new { code = 500, msg = rm.GetString("false") + " " + e.Message }, JsonRequestBehavior.AllowGet);
            }
        }
        [HttpGet]
        public JsonResult confirmedReports(string name, string date)
        {
            try
            {
                string date1 = date == "" ? "" : date.Substring(0, date.IndexOf("/")).Trim();
                string date2 = date == "" ? "" : date.Substring(date.IndexOf("/") + 1).Trim();
                var s1 = sumDate(date1);
                var s2 = sumDate(date2);
                var reports = (from b in db.Reports.Where(x => x.Status == true && x.So == "")
                               select new
                               {
                                   id = b.Id,
                                   createDate = b.ModifyDate,
                                   createBy = b.ModifyBy
                               }).ToList();
                if (name != "" && date != "")
                {
                    if (name == "-1")
                    {
                        var reportss = reports.Where(x => (Convert.ToInt32(x.createDate.Value.Day) + (Convert.ToInt32(x.createDate.Value.Month) * 30) + Convert.ToInt32(x.createDate.Value.Year)) >= s1
                        && (Convert.ToInt32(x.createDate.Value.Day) + (Convert.ToInt32(x.createDate.Value.Month) * 30) + Convert.ToInt32(x.createDate.Value.Year)) <= s2);
                        return Json(new { code = 200, reports = reportss }, JsonRequestBehavior.AllowGet);
                    }
                    else
                    {
                        var user = db.Users.SingleOrDefault(x => x.Id.ToString() == name);
                        name = user.Name;
                        var reportss = reports.Where(x => (Convert.ToInt32(x.createDate.Value.Day) + (Convert.ToInt32(x.createDate.Value.Month) * 30) + Convert.ToInt32(x.createDate.Value.Year)) >= s1 && x.createBy == name
                        && (Convert.ToInt32(x.createDate.Value.Day) + (Convert.ToInt32(x.createDate.Value.Month) * 30) + Convert.ToInt32(x.createDate.Value.Year)) <= s2);
                        return Json(new { code = 200, reports = reportss }, JsonRequestBehavior.AllowGet);
                    }
                }
                return Json(new { code = 200, reports = reports }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception e)
            {
                return Json(new { code = 500, msg = rm.GetString("false") + " " + e.Message }, JsonRequestBehavior.AllowGet);
            }
        }
        [HttpGet]
        public JsonResult FetchEPCDiscrepancyHandheld(string id, int page,string po)
        {
            try
            {
                int pageSize = 1000;
                var epcToUpc = (from a in db.EPCDiscrepancies.Where(x => x.IdReports == id && x.Po == po)
                                orderby a.Id
                                select new
                                {
                                    so = a.So,
                                    po = a.Po,
                                    sku = a.Sku,
                                    EPC = a.Id,
                                    upc = a.UPC,
                                    ctn = a.Carton,
                                    status = a.Status,
                                    a.cntry,
                                    dptPortCd = a.port,
                                    a.doNo,
                                    a.setCd,
                                    a.subDoNo,
                                    a.mngFctryCd,
                                    a.facBranchCd,
                                    a.packKey
                                }).ToList();
                var pages = epcToUpc.Count() % pageSize == 0 ? epcToUpc.Count() / pageSize : epcToUpc.Count() / pageSize + 1;
                epcToUpc = epcToUpc.Skip((page - 1) * pageSize).Take(pageSize).ToList();
                return Json(new
                {
                    code = 200,
                    epcToUpc = epcToUpc,
                    pages
                }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception e)
            {
                return Json(new { code = 500, msg = rm.GetString("false") + " " + e.Message }, JsonRequestBehavior.AllowGet);
            }
        }
        [HttpGet]
        public JsonResult InfoEPCHandheld(string id, int page,string PO)
        {
            try
            {
                int pageSize = 2000;
                var createDates = db.Reports.SingleOrDefault(x => x.Id == id).CreateDate;
                var modifyDates = db.Reports.SingleOrDefault(x => x.Id == id).ModifyDate;
                var createDate = createDates.Value.Day + "-" + createDates.Value.Month + "-" + createDates.Value.Year + " " + createDates.Value.Hour + ":" + createDates.Value.Minute;
                var modifyDate = modifyDates == null ? "" : modifyDates.Value.Day + "-" + modifyDates.Value.Month + "-" + modifyDates.Value.Year + " " + modifyDates.Value.Hour + ":" + modifyDates.Value.Minute;
                var createBy = db.Reports.SingleOrDefault(x => x.Id == id).CreateBy;
                var Consignee = db.Reports.SingleOrDefault(x => x.Id == id).Name;
                var Shipper = db.Reports.SingleOrDefault(x => x.Id == id).Shiper;
                var po = db.Reports.SingleOrDefault(x => x.Id == id).Po;
                var so = db.Reports.SingleOrDefault(x => x.Id == id).So;
                var sku = db.Reports.SingleOrDefault(x => x.Id == id).Sku;
                var epcToUpc = (from a in db.Data.Where(x => x.IdReports == id&&x.Po == PO)
                                orderby a.EPC
                                select new
                                {
                                    so = a.So,
                                    po = a.Po,
                                    sku = a.Sku,
                                    epc = a.EPC,
                                    upc = a.UPC,
                                    ctn = a.CartonTo,
                                    status = a.Status,
                                    a.EPC,
                                    a.cntry,
                                    dptPortCd = a.port,
                                    a.doNo,
                                    a.setCd,
                                    a.subDoNo,
                                    a.mngFctryCd,
                                    a.facBranchCd,
                                    a.packKey
                                }).ToList();
                var pages = epcToUpc.Count() % pageSize == 0 ? epcToUpc.Count() / pageSize : epcToUpc.Count() / pageSize + 1;
                epcToUpc = epcToUpc.Skip((page - 1) * pageSize).Take(pageSize).ToList();
                return Json(new
                {
                    code = 200,
                    createDate = createDate,
                    createBy = createBy,
                    epcToUpc = epcToUpc,
                    po = po,
                    so = so,
                    sku = sku,
                    modifyDate = modifyDate,
                    Consignee = Consignee,
                    Shipper = Shipper,
                    pages
                }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception e)
            {
                return Json(new { code = 500, msg = rm.GetString("false") + " " + e.Message }, JsonRequestBehavior.AllowGet);
            }
        }
        [HttpGet]
        public JsonResult InfoCtn(string id, int page,string PO)
        {
            try
            {
                int pageSize = 1000;
                var createDates = db.Reports.SingleOrDefault(x => x.Id == id).CreateDate;
                var modifyDates = db.Reports.SingleOrDefault(x => x.Id == id).ModifyDate;
                var createDate = createDates.Value.Day + "-" + createDates.Value.Month + "-" + createDates.Value.Year + " " + createDates.Value.Hour + ":" + createDates.Value.Minute;
                var modifyDate = modifyDates == null ? "" : modifyDates.Value.Day + "-" + modifyDates.Value.Month + "-" + modifyDates.Value.Year + " " + modifyDates.Value.Hour + ":" + modifyDates.Value.Minute;
                var createBy = db.Reports.SingleOrDefault(x => x.Id == id).CreateBy;
                var Consignee = db.Reports.SingleOrDefault(x => x.Id == id).Name;
                var Shipper = db.Reports.SingleOrDefault(x => x.Id == id).Shiper;
                var po = db.Reports.SingleOrDefault(x => x.Id == id).Po;
                var so = db.Reports.SingleOrDefault(x => x.Id == id).So;
                var sku = db.Reports.SingleOrDefault(x => x.Id == id).Sku;
                var Ctn = (from a in db.DataScanPhysicals.Where(x => x.IdReports == id&&x.Po == PO)
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
                var pages = Ctn.Count() % pageSize == 0 ? Ctn.Count() / pageSize : Ctn.Count() / pageSize + 1;
                Ctn = Ctn.Skip((page - 1) * pageSize).Take(pageSize).ToList();
                return Json(new
                {
                    code = 200,
                    createDate = createDate,
                    createBy = createBy,
                    Ctn = Ctn,
                    po = po,
                    so = so,
                    sku = sku,
                    modifyDate = modifyDate,
                    Consignee = Consignee,
                    Shipper = Shipper,
                    TotalPages = pages,
                }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception e)
            {
                return Json(new { code = 500, msg = rm.GetString("false") + " " + e.Message }, JsonRequestBehavior.AllowGet);
            }
        }
        public static int sumDate(string date)
        {
            var y = date == "" ? "" : date.Substring(0, date.IndexOf("-"));
            var m = date == "" ? "" : date.Substring(date.IndexOf("-") + 1, 2);
            var d = date == "" ? "" : date.Substring(date.LastIndexOf("-") + 1, 2);
            var s = date == "" ? 0 : int.Parse(d) + (int.Parse(m) * 30) + int.Parse(y);
            return s;
        }
    }
}