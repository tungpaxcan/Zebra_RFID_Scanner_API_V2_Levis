using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data.Entity.Validation;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Resources;
using System.Web;
using System.Web.Hosting;
using System.Web.Mvc;
using System.Web.Services.Description;
using System.Web.UI.WebControls.WebParts;
using System.Windows.Media;
using Zebra_RFID_Scanner.Models;

namespace Zebra_RFID_Scanner.Controllers
{
    public class ReportsController : BaseController
    {
        private Entities db = new Entities();
        ResourceManager rm = new ResourceManager("Zebra_RFID_Scanner.Resources.Resource", typeof(Resources.Resource).Assembly);
        // GET: Reports
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
                               where b.Status == true 
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
        public JsonResult rescanAll(string id)
        {
            try
            {
                string port = "";

                port = checkPort(id);
                if (string.IsNullOrEmpty(port))
                {
                    return Json(new { code = 200, url = "/FunctionOrder/RescanAll",status = false }, JsonRequestBehavior.AllowGet);
                }
                else
                {
                    return Json(new { code = 200, url = "/FunctionOrder/RescanAll2",status = true }, JsonRequestBehavior.AllowGet);
                }
               
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
                var reports = (from b in db.Reports.Where(x => x.Status == false && x.So !="")
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
        public JsonResult UnconfirmedReportsToSearch(string search)
        {
            try
            {
                var reports = (from b in db.Reports.Where(x => x.Status == false)
                               select new
                               {
                                   id = b.Id,
                                   createDate = b.CreateDate,
                                   createBy = b.CreateBy,
                               }).ToList().Where(x=>x.id.Contains(search));
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
                var reports = (from b in db.Reports.Where(x => x.Status == true && x.So != "")
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
        public JsonResult UserReports()
        {
            try
            {
                var reports = (from b in db.Users.Where(x => x.Id > 0)
                               select new
                               {
                                   name = b.Name,
                                   id = b.Id
                               }).ToList();
                return Json(new { code = 200, reports = reports }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception e)
            {
                return Json(new { code = 500, msg = rm.GetString("false") + " " + e.Message }, JsonRequestBehavior.AllowGet);
            }
        }
        [HttpPost]
        public JsonResult Remove(string id)
        {
            try
            {
                Dele.DeleteReports(id);
                return Json(new { code = 200, }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception e)
            {
                return Json(new { code = 500, msg = rm.GetString("false") + " " + e.Message }, JsonRequestBehavior.AllowGet);
            }
        }
        [HttpPost]
        public JsonResult RemoveAll(string id)
        {
            try
            {
                var arrs = JsonConvert.DeserializeObject<string[]>(id);
                foreach (var item in arrs)
                {
                    if (item != null && item != "" && !string.IsNullOrEmpty(item))
                        Dele.DeleteReports(item);
                }
                return Json(new { code = 200, }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception e)
            {
                return Json(new { code = 500, msg = rm.GetString("false") + " " + e.Message }, JsonRequestBehavior.AllowGet);
            }
        }
        [HttpPost]
        public JsonResult reportAlls(string id)
        {
            try
            {
                var arrs = JsonConvert.DeserializeObject<string[]>(id);
                foreach (var item in arrs)
                {
                    if (item != null && item != "" && !string.IsNullOrEmpty(item))
                    {
                        var report = db.Reports.Find(item);
                        if (report != null)
                        {
                            var name = report.So;
                            var hod = report.ModifyDate.ToString().Replace("/", "-").Replace("\\", "-");
                            reportAll(name, hod, item);
                        }

                    }

                }
                return Json(new { code = 200, }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception e)
            {
                return Json(new { code = 500, msg = rm.GetString("false") + " " + e.Message }, JsonRequestBehavior.AllowGet);
            }
        }
        [HttpGet]
        public JsonResult InfoReports(string id)
        {
            try
            {
                var createDates = db.Reports.SingleOrDefault(x => x.Id == id).CreateDate;
                var modifyDates = db.Reports.SingleOrDefault(x => x.Id == id).ModifyDate;
                var timeStarts = db.Reports.SingleOrDefault(x => x.Id == id).TimeStart == null ? DateTime.Now : db.Reports.SingleOrDefault(x => x.Id == id).TimeStart;
                var createDate = createDates.Value.Day + "-" + createDates.Value.Month + "-" + createDates.Value.Year + " " + createDates.Value.Hour + ":" + createDates.Value.Minute + ":" + createDates.Value.Second;
                var modifyDate = modifyDates == null ? "" : modifyDates.Value.Day + "-" + modifyDates.Value.Month + "-" + modifyDates.Value.Year + " " + modifyDates.Value.Hour + ":" + modifyDates.Value.Minute;
                var timeStart = timeStarts.Value.Day + "-" + timeStarts.Value.Month + "-" + timeStarts.Value.Year + " " + timeStarts.Value.Hour + ":" + timeStarts.Value.Minute + ":" + timeStarts.Value.Second;
                var createBy = db.Reports.SingleOrDefault(x => x.Id == id).CreateBy;
                var Consignee = db.Reports.SingleOrDefault(x => x.Id == id).Name;
                var Shipper = db.Reports.SingleOrDefault(x => x.Id == id).Shiper;
                var po = db.Reports.SingleOrDefault(x => x.Id == id).Po;
                var so = db.Reports.SingleOrDefault(x => x.Id == id).So;
                var sku = db.Reports.SingleOrDefault(x => x.Id == id).Sku;
                return Json(new
                {
                    code = 200,
                    createDate = createDate,
                    createBy = createBy,
                    po = po,
                    so = so,
                    sku = sku,
                    modifyDate = modifyDate,
                    Consignee = Consignee,
                    Shipper = Shipper,
                    timeStart = timeStart
                }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception e)
            {
                return Json(new { code = 500, msg = rm.GetString("false") + " " + e.Message }, JsonRequestBehavior.AllowGet);
            }
        }
        [HttpGet]
        public JsonResult InfoCtn(string id, int page)
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
                var Ctn = (from a in db.DataScanPhysicals.Where(x => x.IdReports == id)
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
        [HttpGet]
        public JsonResult InfoDis(string id, int page)
        {
            try
            {
                string port = "";
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
                port = checkPort(id);
                var discrepancies = (from a in db.Discrepancies.Where(x => x.IdReports == id)
                                     select new
                                     {
                                         id = a.Id,
                                         so = a.So,
                                         po = a.Po,
                                         sku = a.Sku,
                                         ctn = a.CartonTo,
                                         upc = a.UPC,
                                         qty = a.Qty,
                                         qtyScan = a.QtyScan
                                     }).ToList();
                var totalPages = discrepancies.Count / pageSize + 1;
                var general = (from a in db.Generals.Where(x => x.IdReports == id)
                               select new
                               {
                                   so = a.So,
                                   po = a.Po,
                                   sku = a.Sku,
                                   ctn = a.CartonTo,
                                   upc = a.UPC,
                                   qty = a.Qty,
                               }).ToList();
                var pages = discrepancies.Count() % pageSize == 0 ? discrepancies.Count() / pageSize : discrepancies.Count() / pageSize + 1;
                discrepancies = discrepancies.Skip((page - 1) * pageSize).Take(pageSize).ToList();
                return Json(new
                {
                    code = 200,
                    createDate = createDate,
                    createBy = createBy,
                    discrepancies = discrepancies,
                    general = general,
                    po = po,
                    so = so,
                    sku = sku,
                    modifyDate = modifyDate,
                    Consignee = Consignee,
                    Shipper = Shipper,
                    totalPages = totalPages,
                    fileCsv= port==null?"":port.ToString(),
                }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception e)
            {
                return Json(new { code = 500, msg = rm.GetString("false") + " " + e.Message }, JsonRequestBehavior.AllowGet);
            }
        }
        [HttpGet]
        public JsonResult InfoGeneral(string id, int page)
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
                var general = (from a in db.Generals.Where(x => x.IdReports == id)
                               select new
                               {
                                   so = a.So,
                                   po = a.Po,
                                   sku = a.Sku,
                                   ctn = a.CartonTo,
                                   upc = a.UPC,
                                   qty = a.Qty,
                               }).ToList();
                var pages = general.Count() % pageSize == 0 ? general.Count() / pageSize : general.Count() / pageSize + 1;
                general = general.Skip((page - 1) * pageSize).Take(pageSize).ToList();
                return Json(new
                {
                    code = 200,
                    createDate = createDate,
                    createBy = createBy,
                    general = general,
                    po = po,
                    so = so,
                    sku = sku,
                    modifyDate = modifyDate,
                    Consignee = Consignee,
                    Shipper = Shipper
                }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception e)
            {
                return Json(new { code = 500, msg = rm.GetString("false") + " " + e.Message }, JsonRequestBehavior.AllowGet);
            }
        }
        [HttpGet]
        public JsonResult FetchEPCDiscrepancy(string id, int page)
        {
            try
            {
                int pageSize = 1000;
                var epcToUpc = (from a in db.EPCDiscrepancies.Where(x => x.IdReports == id)
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
        public JsonResult InfoEPC(string id, int page)
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
                var epcToUpc = (from a in db.Data.Where(x => x.IdReports == id)
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
        public JsonResult confirmHod(string id)
        {
            try
            {
                var user = (User)Session["user"];
                var name = user.Name;
                var date = DateTime.Now.Day + "/" + DateTime.Now.Month + "/" + DateTime.Now.Year + " " + DateTime.Now.Hour + ":" + DateTime.Now.Minute;
                var reports = db.Reports.SingleOrDefault(x => x.Id == id);
                reports.Status = true;
                reports.ModifyDate = DateTime.Now;
                reports.ModifyBy = name;
                db.SaveChanges();
                return Json(new { code = 200, date = date, id = id }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception e)
            {
                return Json(new { code = 500, msg = "Sai !!!" + e.Message }, JsonRequestBehavior.AllowGet);
            }
        }
        [HttpPost]
        public JsonResult reportSummary(string date)
        {
            try
            {
                string date1 = date == "" ? "" : date.Substring(0, date.IndexOf("/")).Trim();
                string date2 = date == "" ? "" : date.Substring(date.IndexOf("/") + 1).Trim();
                var s1 = sumDate(date1);
                var s2 = sumDate(date2);
                List<ReportsEPC> datas = new List<ReportsEPC>();
                ReportsEPC data = new ReportsEPC()
                {
                    C1 = "CONSIGNEE",
                    C2 = "SHIPPER",
                    C3 = "SO NUMBER",
                    C4 = "PO NUMBER",
                    C5 = "SKU",
                    C6 = "TOTAL UNITS",
                    C7 = "SCANNING UNITS",
                    C8 = "VARIANCE",
                    C9 = "RESULT",
                    C10 = "START SCANNING DATE TIME",
                    C11 = "END SCANNING TIME",
                    C12 = "REMARK",
                    C13 = "HOD"
                };
                datas.Add(data);
                var reports = db.Reports.Where(x => x.Id.Length > 0 && (x.CreateDate.Value.Day + (x.CreateDate.Value.Month * 30) + x.CreateDate.Value.Year) >= s1
                        && (x.CreateDate.Value.Day + (x.CreateDate.Value.Month * 30) + x.CreateDate.Value.Year) <= s2).ToList();
                foreach (var item in reports)
                {
                    var qty = TotalQTYReports(item.Id).Qty;
                    var qtyScan = TotalQTYReports(item.Id).QtyScan;
                    var qtyCtn = TotalQTYReports(item.Id).QtyCtn;
                    ReportsEPC datum = new ReportsEPC()
                    {
                        C1 = item.Name,
                        C2 = item.Shiper,
                        C3 = item.So,
                        C4 = item.Po,
                        C5 = item.Sku,
                        C6 = qty.ToString(),
                        C7 = qtyScan.ToString(),
                        C8 = (qtyScan - qty).ToString(),
                        C9 = (qtyScan - qty) == 0 ? "Matched" : "Mismatched",
                        C10 = item.TimeStart == null ? "" : item.TimeStart.Value.Month + "-" + item.TimeStart.Value.Day + "-" + item.TimeStart.Value.Year + " " + item.TimeStart.Value.Hour + ":" + item.TimeStart.Value.Minute + ":" + item.TimeStart.Value.Second,
                        C11 = item.CreateDate == null ? "" : item.CreateDate.Value.Month + "-" + item.CreateDate.Value.Day + "-" + item.CreateDate.Value.Year + " " + item.CreateDate.Value.Hour + ":" + item.CreateDate.Value.Minute + ":" + item.CreateDate.Value.Second,
                        C12 = "",
                        C13 = item.ModifyDate == null ? "" : item.ModifyDate.Value.Month + "-" + item.ModifyDate.Value.Day + "-" + item.ModifyDate.Value.Year + " " + item.ModifyDate.Value.Hour + ":" + item.ModifyDate.Value.Minute + ":" + item.ModifyDate.Value.Second,
                    };
                    datas.Add(datum);
                }
                string folderName = @"C:\RFID\Reports\";
                string filePath = "MRA RFID Scaning Report.xlsx";
                string folderPath = folderName + filePath;
                using (ExcelPackage excelPackage = new ExcelPackage())
                {
                    // Add a new worksheet to the workbook
                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");

                    // Add some content to the worksheet
                    for (int i = 0; i < datas.Count; i++)
                    {
                        for (int j = 0; j < 13; j++)
                        {
                            var propertyName = "C" + (j + 1);
                            var property = typeof(ReportsEPC).GetProperty(propertyName);
                            var propertyValue = property.GetValue(datas[i]);
                            worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Value = propertyValue;
                            worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                            worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Top.Color.SetColor(System.Drawing.Color.Black);
                            worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Bottom.Color.SetColor(System.Drawing.Color.Black);
                            worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                            worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Left.Color.SetColor(System.Drawing.Color.Black);
                            worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                            worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Right.Color.SetColor(System.Drawing.Color.Black);
                            if (i == 0) // Nếu là hàng đầu tiên
                            {
                                worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 255, 0));
                            }
                        }
                    }
                    worksheet.View.FreezePanes(2, 1);
                    // Save the ExcelPackage to a file
                    FileInfo file = new FileInfo(folderPath);
                    excelPackage.SaveAs(file);
                }
                return Json(new { code = 200, }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception e)
            {
                return Json(new { code = 500, msg = " !!!" + e.InnerException }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public JsonResult reportGeneral(string id, string name, string hod)
        {
            try
            {
                hod = hod.Substring(0, hod.IndexOf(" "));
                List<ReportsEPC> datas = new List<ReportsEPC>();
                ReportsEPC data = new ReportsEPC()
                {
                    C1 = "CONSIGNEE",
                    C2 = "SHIPPER",
                    C3 = "SO NUMBER",
                    C4 = "PO NUMBER",
                    C5 = "SKU",
                    C6 = "CARTONID",
                    C7 = "UPC",
                    C8 = "TOTAL UNITS",
                    C9 = "SCANNING UNITS",
                    C10 = "RESULT",
                    C11 = "SCANNING DATE TIME",
                    C12 = "REMARK",
                    C13 = "HOD"
                };
                datas.Add(data);
                var reports = db.Reports.SingleOrDefault(x => x.Id == id);
                var list = db.Generals.Where(x => x.IdReports == id);

                foreach (var item in list)
                {
                    ReportsEPC datum = new ReportsEPC()
                    {
                        C1 = reports.Name,
                        C2 = reports.Shiper,
                        C3 = item.So,
                        C4 = item.Po,
                        C5 = item.Sku,
                        C6 = item.CartonTo,
                        C7 = item.UPC,
                        C8 = item.Qty,
                        C9 = item.Qty,
                        C10 = "Matched",
                        C11 = item.CreateDate == null ? "" : item.CreateDate.Value.Month + "-" + item.CreateDate.Value.Day + "-" + item.CreateDate.Value.Year + " " + item.CreateDate.Value.Hour + ":" + item.CreateDate.Value.Minute + ":" + item.CreateDate.Value.Second,
                        C12 = "",
                        C13 = reports.ModifyDate == null ? "" : reports.ModifyDate.Value.Month + "-" + reports.ModifyDate.Value.Day + "-" + reports.ModifyDate.Value.Year + " " + reports.ModifyDate.Value.Hour + ":" + reports.ModifyDate.Value.Minute + ":" + reports.ModifyDate.Value.Second,
                    };
                    datas.Add(datum);
                }
                string folderName = @"C:\RFID\Reports\";
                string filePath = "General" + "_" + name + "_" + hod + ".xlsx";
                string folderPath = folderName + filePath;
                using (ExcelPackage excelPackage = new ExcelPackage())
                {
                    // Add a new worksheet to the workbook
                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");

                    // Add some content to the worksheet
                    for (int i = 0; i < datas.Count; i++)
                    {
                        for (int j = 0; j < 13; j++)
                        {
                            var propertyName = "C" + (j + 1);
                            var property = typeof(ReportsEPC).GetProperty(propertyName);
                            var propertyValue = property.GetValue(datas[i]);
                            worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Value = propertyValue;
                            worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                            worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Top.Color.SetColor(System.Drawing.Color.Black);
                            worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Bottom.Color.SetColor(System.Drawing.Color.Black);
                            worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                            worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Left.Color.SetColor(System.Drawing.Color.Black);
                            worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                            worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Right.Color.SetColor(System.Drawing.Color.Black);
                            if (i == 0) // Nếu là hàng đầu tiên
                            {
                                worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 255, 0));
                            }
                        }
                    }
                    worksheet.View.FreezePanes(2, 1);
                    // Save the ExcelPackage to a file
                    FileInfo file = new FileInfo(folderPath);
                    excelPackage.SaveAs(file);
                }
                return Json(new { code = 200, }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception e)
            {
                return Json(new { code = 500, msg = " !!!" + e.Message }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public JsonResult reportCartonScan(string[] data, string name, string hod)
        {
            try
            {
                hod = hod.Substring(0, hod.IndexOf(" "));
                Report(data, name, hod, "CartonScan");
                return Json(new { code = 200, }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception e)
            {
                return Json(new { code = 500, msg = " !!!" + e.Message }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public JsonResult reportDiscrepancy(string id, string name, string hod)
        {
            try
            {
                hod = hod.Substring(0, hod.IndexOf(" "));
                List<ReportsEPC> datas = new List<ReportsEPC>();
                ReportsEPC data = new ReportsEPC()
                {
                    C1 = "CONSIGNEE",
                    C2 = "SHIPPER",
                    C3 = "SO NUMBER",
                    C4 = "PO NUMBER",
                    C5 = "SKU",
                    C6 = "CARTONID",
                    C7 = "UPC",
                    C8 = "TOTAL UNITS",
                    C9 = "SCANNING UNITS",
                    C10 = "RESULT",
                    C11 = "SCANNING DATE TIME",
                    C12 = "REMARK",
                    C13 = "HOD"
                };
                datas.Add(data);
                var reports = db.Reports.SingleOrDefault(x => x.Id == id);
                var list = db.Discrepancies.Where(x => x.IdReports == id);

                foreach (var item in list)
                {
                    ReportsEPC datum = new ReportsEPC()
                    {
                        C1 = reports.Name,
                        C2 = reports.Shiper,
                        C3 = item.So,
                        C4 = item.Po,
                        C5 = item.Sku,
                        C6 = item.CartonTo,
                        C7 = item.UPC,
                        C8 = item.Qty,
                        C9 = item.QtyScan,
                        C10 = "Mismatched",
                        C11 = item.CreateDate == null ? "" : item.CreateDate.Value.Month + "-" + item.CreateDate.Value.Day + "-" + item.CreateDate.Value.Year + " " + item.CreateDate.Value.Hour + ":" + item.CreateDate.Value.Minute + ":" + item.CreateDate.Value.Second,
                        C12 = "",
                        C13 = reports.ModifyDate == null ? "" : reports.ModifyDate.Value.Month + "-" + reports.ModifyDate.Value.Day + "-" + reports.ModifyDate.Value.Year + " " + reports.ModifyDate.Value.Hour + ":" + reports.ModifyDate.Value.Minute + ":" + reports.ModifyDate.Value.Second,
                    };
                    datas.Add(datum);
                }
                string folderName = @"C:\RFID\Reports\";
                string filePath = "Discrepancy" + "_" + name + "_" + hod + ".xlsx";
                string folderPath = folderName + filePath;
                using (ExcelPackage excelPackage = new ExcelPackage())
                {
                    // Add a new worksheet to the workbook
                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");

                    // Add some content to the worksheet
                    for (int i = 0; i < datas.Count; i++)
                    {
                        for (int j = 0; j < 13; j++)
                        {
                            var propertyName = "C" + (j + 1);
                            var property = typeof(ReportsEPC).GetProperty(propertyName);
                            var propertyValue = property.GetValue(datas[i]);
                            worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Value = propertyValue;
                            worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                            worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Top.Color.SetColor(System.Drawing.Color.Black);
                            worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Bottom.Color.SetColor(System.Drawing.Color.Black);
                            worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                            worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Left.Color.SetColor(System.Drawing.Color.Black);
                            worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                            worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Right.Color.SetColor(System.Drawing.Color.Black);
                            if (i == 0) // Nếu là hàng đầu tiên
                            {
                                worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 255, 0));
                            }
                        }
                    }
                    worksheet.View.FreezePanes(2, 1);
                    // Save the ExcelPackage to a file
                    FileInfo file = new FileInfo(folderPath);
                    excelPackage.SaveAs(file);
                }
                return Json(new { code = 200, }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception e)
            {
                return Json(new { code = 500, msg = " !!!" + e.Message }, JsonRequestBehavior.AllowGet);
            }
        }
        [HttpPost]
        public JsonResult reportDiscrepancyTotal(string[] data, string name, string hod)
        {
            try
            {
                hod = hod.Substring(0, hod.IndexOf(" "));
                Report(data, name, hod, "DiscrepancyTotal");
                return Json(new { code = 200, }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception e)
            {
                return Json(new { code = 500, msg = " !!!" + e.Message }, JsonRequestBehavior.AllowGet);
            }
        }
        [HttpPost]
        public JsonResult reportEPC(string id, string name, string hod)
        {
            try
            {
                hod = hod.Substring(0, hod.IndexOf(" "));
                List<ReportsEPC> datas = new List<ReportsEPC>();
                ReportsEPC data = new ReportsEPC()
                {
                    C1 = "CONSIGNEE",
                    C2 = "SHIPPER",
                    C3 = "OPERATOR",
                    C4 = "SO",
                    C5 = "PO",
                    C6 = "SKU",
                    C7 = "EPC",
                    C8 = "UPC",
                    C9 = "SCANNING DATE TIME",
                    C10 = "HOD"
                };
                datas.Add(data);
                var reports = db.Reports.SingleOrDefault(x => x.Id == id);
                var list = db.Data.Where(x => x.IdReports == id);

                foreach (var item in list)
                {
                    ReportsEPC datum = new ReportsEPC()
                    {
                        C1 = reports.Name,
                        C2 = reports.Shiper,
                        C3 = reports.CreateBy,
                        C4 = item.So,
                        C5 = item.Po,
                        C6 = item.Sku,
                        C7 = item.EPC,
                        C8 = item.UPC,
                        C9 = item.CreateDate == null ? "" : item.CreateDate.Value.Month + "-" + item.CreateDate.Value.Day + "-" + item.CreateDate.Value.Year + " " + item.CreateDate.Value.Hour + ":" + item.CreateDate.Value.Minute + ":" + item.CreateDate.Value.Second,
                        C10 = reports.ModifyDate == null ? "" : reports.ModifyDate.Value.Month + "-" + reports.ModifyDate.Value.Day + "-" + reports.ModifyDate.Value.Year + " " + reports.ModifyDate.Value.Hour + ":" + reports.ModifyDate.Value.Minute + ":" + reports.ModifyDate.Value.Second,
                    };
                    datas.Add(datum);
                }

                string folderName = @"C:\RFID\Reports\";
                string filePath = "EPC" + "_" + name + "_" + hod + ".xlsx";
                string folderPath = folderName + filePath;
                using (ExcelPackage excelPackage = new ExcelPackage())
                {
                    // Add a new worksheet to the workbook
                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");

                    // Add some content to the worksheet
                    for (int i = 0; i < datas.Count; i++)
                    {
                        for (int j = 0; j < 10; j++)
                        {
                            var propertyName = "C" + (j + 1);
                            var property = typeof(ReportsEPC).GetProperty(propertyName);
                            var propertyValue = property.GetValue(datas[i]);
                            worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Value = propertyValue;
                            worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                            worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Top.Color.SetColor(System.Drawing.Color.Black);
                            worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Bottom.Color.SetColor(System.Drawing.Color.Black);
                            worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                            worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Left.Color.SetColor(System.Drawing.Color.Black);
                            worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                            worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Right.Color.SetColor(System.Drawing.Color.Black);
                            if (i == 0) // Nếu là hàng đầu tiên
                            {
                                worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 255, 0));
                            }
                        }
                    }
                    worksheet.View.FreezePanes(2, 1);
                    // Save the ExcelPackage to a file
                    FileInfo file = new FileInfo(folderPath);
                    excelPackage.SaveAs(file);
                }
                return Json(new { code = 200, }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception e)
            {
                return Json(new { code = 500, msg = "Sai !!!" + e.Message }, JsonRequestBehavior.AllowGet);
            }
        }
        [HttpPost]
        public JsonResult reportAll(string name, string hod, string id)
        {
            try
            {
                hod = hod.Substring(0, hod.IndexOf(" "));
                string folderName = @"C:\RFID\Reports\";
                string filePath = "Reports All " + name + "_" + hod + ".xlsx";
                string folderPath = folderName + filePath;
                var reports = db.Reports.SingleOrDefault(x => x.Id == id);
                using (ExcelPackage excelPackage = new ExcelPackage())
                {
                    List<ReportsEPC> datasGen = new List<ReportsEPC>();
                    ReportsEPC dataGen = new ReportsEPC()
                    {
                        C1 = "CONSIGNEE",
                        C2 = "SHIPPER",
                        C3 = "SO NUMBER",
                        C4 = "PO NUMBER",
                        C5 = "SKU",
                        C6 = "CARTONID",
                        C7 = "UPC",
                        C8 = "TOTAL UNITS",
                        C9 = "SCANNING UNITS",
                        C10 = "RESULT",
                        C11 = "START SCANNING DATE TIME",
                        C12 = "END SCANNING TIME",
                        C13 = "REMARK",
                        C14 = "HOD"
                    };
                    datasGen.Add(dataGen);
                    var listGen = db.Generals.Where(x => x.IdReports == id);
                    foreach (var item in listGen)
                    {
                        ReportsEPC datum = new ReportsEPC()
                        {
                            C1 = reports.Name,
                            C2 = reports.Shiper,
                            C3 = item.So,
                            C4 = item.Po,
                            C5 = item.Sku,
                            C6 = item.CartonTo,
                            C7 = item.UPC,
                            C8 = item.Qty,
                            C9 = item.Qty,
                            C10 = "Matched",
                            C11 = item.Report.TimeStart == null ? "" : item.Report.TimeStart.Value.Month + "-" + item.Report.TimeStart.Value.Day + "-" + item.Report.TimeStart.Value.Year + " " + item.Report.TimeStart.Value.Hour + ":" + item.Report.TimeStart.Value.Minute + ":" + item.Report.TimeStart.Value.Second,
                            C12 = item.CreateDate == null ? "" : item.CreateDate.Value.Month + "-" + item.CreateDate.Value.Day + "-" + item.CreateDate.Value.Year + " " + item.CreateDate.Value.Hour + ":" + item.CreateDate.Value.Minute + ":" + item.CreateDate.Value.Second,
                            C13 = "",
                            C14 = reports.ModifyDate == null ? "" : reports.ModifyDate.Value.Month + "-" + reports.ModifyDate.Value.Day + "-" + reports.ModifyDate.Value.Year + " " + reports.ModifyDate.Value.Hour + ":" + reports.ModifyDate.Value.Minute + ":" + reports.ModifyDate.Value.Second,
                        };
                        datasGen.Add(datum);
                    }
                    // Add a new worksheet to the workbook
                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("General");

                    // Add some content to the worksheet
                    for (int i = 0; i < datasGen.Count; i++)
                    {
                        for (int j = 0; j < 14; j++)
                        {
                            var propertyName = "C" + (j + 1);
                            var property = typeof(ReportsEPC).GetProperty(propertyName);
                            var propertyValue = property.GetValue(datasGen[i]);
                            worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Value = propertyValue;
                            worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                            worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Top.Color.SetColor(System.Drawing.Color.Black);
                            worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Bottom.Color.SetColor(System.Drawing.Color.Black);
                            worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                            worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Left.Color.SetColor(System.Drawing.Color.Black);
                            worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                            worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Right.Color.SetColor(System.Drawing.Color.Black);
                            if (i == 0) // Nếu là hàng đầu tiên
                            {
                                worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 255, 0));
                            }
                        }
                    }
                    worksheet.View.FreezePanes(2, 1);
                    List<ReportsEPC> datasDis = new List<ReportsEPC>();
                    ReportsEPC dataDis = new ReportsEPC()
                    {
                        C1 = "CONSIGNEE",
                        C2 = "SHIPPER",
                        C3 = "SO NUMBER",
                        C4 = "PO NUMBER",
                        C5 = "SKU",
                        C6 = "CARTONID",
                        C7 = "UPC",
                        C8 = "TOTAL UNITS",
                        C9 = "SCANNING UNITS",
                        C10 = "RESULT",
                        C11 = "START SCANNING DATE TIME",
                        C12 = "END SCANNING TIME",
                        C13 = "REMARK",
                        C14 = "HOD"
                    };
                    datasDis.Add(dataDis);

                    var listDis = db.Discrepancies.Where(x => x.IdReports == id);

                    foreach (var item in listDis)
                    {
                        ReportsEPC datum = new ReportsEPC()
                        {
                            C1 = reports.Name,
                            C2 = reports.Shiper,
                            C3 = item.So,
                            C4 = item.Po,
                            C5 = item.Sku,
                            C6 = item.CartonTo,
                            C7 = item.UPC,
                            C8 = item.Qty,
                            C9 = item.QtyScan,
                            C10 = "Mismatched",
                            C11 = item.Report.TimeStart == null ? "" : item.Report.TimeStart.Value.Month + "-" + item.Report.TimeStart.Value.Day + "-" + item.Report.TimeStart.Value.Year + " " + item.Report.TimeStart.Value.Hour + ":" + item.Report.TimeStart.Value.Minute + ":" + item.Report.TimeStart.Value.Second,
                            C12 = item.CreateDate == null ? "" : item.CreateDate.Value.Month + "-" + item.CreateDate.Value.Day + "-" + item.CreateDate.Value.Year + " " + item.CreateDate.Value.Hour + ":" + item.CreateDate.Value.Minute + ":" + item.CreateDate.Value.Second,
                            C13 = "",
                            C14 = reports.ModifyDate == null ? "" : reports.ModifyDate.Value.Month + "-" + reports.ModifyDate.Value.Day + "-" + reports.ModifyDate.Value.Year + " " + reports.ModifyDate.Value.Hour + ":" + reports.ModifyDate.Value.Minute + ":" + reports.ModifyDate.Value.Second,
                        };
                        datasDis.Add(datum);
                    }
                    ExcelWorksheet worksheet2 = excelPackage.Workbook.Worksheets.Add("Discrepancy");

                    // Add some content to the worksheet
                    for (int i = 0; i < datasDis.Count; i++)
                    {
                        for (int j = 0; j < 14; j++)
                        {
                            var propertyName = "C" + (j + 1);
                            var property = typeof(ReportsEPC).GetProperty(propertyName);
                            var propertyValue = property.GetValue(datasDis[i]);
                            worksheet2.Cells[GetExcelColumnName(j + 1) + (i + 1)].Value = propertyValue;
                            worksheet2.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                            worksheet2.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Top.Color.SetColor(System.Drawing.Color.Black);
                            worksheet2.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            worksheet2.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Bottom.Color.SetColor(System.Drawing.Color.Black);
                            worksheet2.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                            worksheet2.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Left.Color.SetColor(System.Drawing.Color.Black);
                            worksheet2.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                            worksheet2.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Right.Color.SetColor(System.Drawing.Color.Black);
                            if (i == 0) // Nếu là hàng đầu tiên
                            {
                                worksheet2.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                worksheet2.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 255, 0));
                            }
                        }
                    }
                    worksheet2.View.FreezePanes(2, 1);
                    List<ReportsEPC> datas = new List<ReportsEPC>();
                    ReportsEPC data = new ReportsEPC()
                    {
                        C1 = "CONSIGNEE",
                        C2 = "SHIPPER",
                        C3 = "OPERATOR",
                        C4 = "SO",
                        C5 = "PO",
                        C6 = "SKU",
                        C7 = "EPC",
                        C8 = "UPC",
                        C9 = "START SCANNING DATE TIME",
                        C10 = "END SCANNING TIME",
                        C11 = "HOD"
                    };
                    datas.Add(data);
                    var list = db.Data.Where(x => x.IdReports == id);

                    foreach (var item in list)
                    {
                        var startTime = db.Reports.SingleOrDefault(x => x.Id == item.IdReports).TimeStart;
                        ReportsEPC datum = new ReportsEPC()
                        {
                            C1 = reports.Name,
                            C2 = reports.Shiper,
                            C3 = reports.CreateBy,
                            C4 = item.So,
                            C5 = item.Po,
                            C6 = item.Sku,
                            C7 = item.EPC,
                            C8 = item.UPC,
                            C9 = startTime == null ? "" : startTime.Value.Month + "-" + startTime.Value.Day + "-" + startTime.Value.Year + " " + startTime.Value.Hour + ":" + startTime.Value.Minute + ":" + startTime.Value.Second,
                            C10 = item.CreateDate == null ? "" : item.CreateDate.Value.Month + "-" + item.CreateDate.Value.Day + "-" + item.CreateDate.Value.Year + " " + item.CreateDate.Value.Hour + ":" + item.CreateDate.Value.Minute + ":" + item.CreateDate.Value.Second,
                            C11 = reports.ModifyDate == null ? "" : reports.ModifyDate.Value.Month + "-" + reports.ModifyDate.Value.Day + "-" + reports.ModifyDate.Value.Year + " " + reports.ModifyDate.Value.Hour + ":" + reports.ModifyDate.Value.Minute + ":" + reports.ModifyDate.Value.Second,
                        };
                        datas.Add(datum);
                    }
                    ExcelWorksheet worksheet3 = excelPackage.Workbook.Worksheets.Add("Epc-Upc");

                    // Add some content to the worksheet
                    for (int i = 0; i < datas.Count; i++)
                    {
                        for (int j = 0; j < 11; j++)
                        {
                            var propertyName = "C" + (j + 1);
                            var property = typeof(ReportsEPC).GetProperty(propertyName);
                            var propertyValue = property.GetValue(datas[i]);
                            worksheet3.Cells[GetExcelColumnName(j + 1) + (i + 1)].Value = propertyValue;
                            worksheet3.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                            worksheet3.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Top.Color.SetColor(System.Drawing.Color.Black);
                            worksheet3.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            worksheet3.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Bottom.Color.SetColor(System.Drawing.Color.Black);
                            worksheet3.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                            worksheet3.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Left.Color.SetColor(System.Drawing.Color.Black);
                            worksheet3.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                            worksheet3.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Right.Color.SetColor(System.Drawing.Color.Black);
                            if (i == 0) // Nếu là hàng đầu tiên
                            {
                                worksheet3.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                worksheet3.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 255, 0));
                            }
                        }
                    }
                    worksheet3.View.FreezePanes(2, 1);
                    // Save the ExcelPackage to a file
                    FileInfo file = new FileInfo(folderPath);
                    excelPackage.SaveAs(file);
                }
                return Json(new { code = 200, }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception e)
            {
                return Json(new { code = 500, msg = "Sai !!!" + e.InnerException.Message }, JsonRequestBehavior.AllowGet);
            }
        }
        [HttpPost]
        public JsonResult reportNew(string id, string name, string hod)
        {
            try
            {
                hod = hod.Substring(0, hod.IndexOf(" "));
                List<ReportsEPC> datas = new List<ReportsEPC>();
                ReportsEPC data = new ReportsEPC()
                {
                    C1 = "Year",
                    C2 = "Fiscal Qtr",
                    C3 = "Scanning Month",
                    C4 = "Week Number",
                    C5 = "Scanning Date",
                    C6 = "Place Of Receipt Country",
                    C7 = "Place Of Receipt City",
                    C8 = "Place Of Delivery",
                    C9 = "Booking Number",
                    C10 = "PO Number",
                    C11 = "SKU ==> Product code:",
                    C12 = "Shipper Code",
                    C13 = "Shipper Name",
                    C14 = "CASE NO.",
                    C15 = "Total Units",
                    C16 = "Read Units",
                    C17 = "Scan Result - Formula",
                    C18 = "Open box?",
                    C19 = "Open Box Counting Result",
                    C20 = "Reason Code of Mismatched cases",
                    C21 = "Remarks",
                };
                datas.Add(data);
                var reports = db.Reports.SingleOrDefault(x => x.Id == id);
                var discrepancies = db.Discrepancies.Where(x => x.IdReports == id);
                var generals = db.Generals.Where(x => x.IdReports == id);
                CultureInfo cultureInfo = CultureInfo.CurrentCulture;
                System.Globalization.Calendar calendar = cultureInfo.Calendar;
                CalendarWeekRule weekRule = cultureInfo.DateTimeFormat.CalendarWeekRule;
                DayOfWeek firstDayOfWeek = cultureInfo.DateTimeFormat.FirstDayOfWeek;
                foreach (var item in discrepancies)
                {
                    string monthName = new DateTime(2000, reports.ModifyDate.Value.Month, 1)
                     .ToString("MMMM", CultureInfo.InvariantCulture);

                    ReportsEPC datum = new ReportsEPC()
                    {
                        C1 = reports.ModifyDate == null ? "" : reports.ModifyDate.Value.Year.ToString(),
                        C2 = reports.ModifyDate == null ? "" : "Q" + Quy(reports.ModifyDate.Value.Month).ToString() + reports.ModifyDate.Value.Year.ToString(),
                        C3 = reports.ModifyDate == null ? "" : monthName,
                        C4 = reports.ModifyDate == null ? "" : calendar.GetWeekOfYear(reports.ModifyDate.Value, weekRule, firstDayOfWeek).ToString(),
                        C5 = reports.CreateDate == null ? "" : reports.CreateDate.Value.Date.ToString(),
                        C6 = reports.Shiper == null ? "" : reports.Shiper.Substring(0, reports.Shiper.IndexOf(" ")),
                        C7 = reports.Shiper == null ? "" : reports.Shiper.Substring(0, reports.Shiper.IndexOf(" ")),
                        C8 = reports.Name == null ? "" : reports.Name.Substring(0, reports.Name.IndexOf(" ")),
                        C9 = item.So == null ? "" : item.So,
                        C10 = item.Po == null ? "" : item.Po,
                        C11 = item.Sku == null ? "" : item.Sku,
                        C12 = reports.Shiper == null ? "" : reports.Shiper,
                        C13 = reports.Id == null ? "" : reports.Id,
                        C14 = item.CartonTo == null ? "" : item.CartonTo,
                        C15 = item.Qty == null ? "" : item.Qty,
                        C16 = item.QtyScan == null ? "" : item.QtyScan,
                        C17 = "Mismatched",
                        C18 = "YES",
                        C19 = (int.Parse(item.Qty) - int.Parse(item.QtyScan)).ToString(),
                        C20 = "Damaged RFID tags",
                        C21 = "",
                    };
                    datas.Add(datum);
                }
                foreach (var item in generals)
                {
                    string monthName = new DateTime(2000, reports.ModifyDate.Value.Month, 1)
                     .ToString("MMMM", CultureInfo.InvariantCulture);

                    ReportsEPC datum = new ReportsEPC()
                    {
                        C1 = reports.ModifyDate == null ? "" : reports.ModifyDate.Value.Year.ToString(),
                        C2 = reports.ModifyDate == null ? "" : "Q" + Quy(reports.ModifyDate.Value.Month).ToString() + reports.ModifyDate.Value.Year.ToString(),
                        C3 = reports.ModifyDate == null ? "" : monthName,
                        C4 = reports.ModifyDate == null ? "" : calendar.GetWeekOfYear(reports.ModifyDate.Value, weekRule, firstDayOfWeek).ToString(),
                        C5 = reports.CreateDate == null ? "" : reports.CreateDate.Value.Date.ToString(),
                        C6 = reports.Shiper == null ? "" : reports.Shiper.Substring(0, reports.Shiper.IndexOf(" ")),
                        C7 = reports.Shiper == null ? "" : reports.Shiper.Substring(0, reports.Shiper.IndexOf(" ")),
                        C8 = reports.Name == null ? "" : reports.Name.Substring(0, reports.Name.IndexOf(" ")),
                        C9 = item.So == null ? "" : item.So,
                        C10 = item.Po == null ? "" : item.Po,
                        C11 = item.Sku == null ? "" : item.Sku,
                        C12 = reports.Shiper == null ? "" : reports.Shiper,
                        C13 = reports.Id == null ? "" : reports.Id,
                        C14 = item.CartonTo == null ? "" : item.CartonTo,
                        C15 = item.Qty == null ? "" : item.Qty,
                        C16 = item.Qty == null ? "" : item.Qty,
                        C17 = "Matched",
                        C18 = "",
                        C19 = "",
                        C20 = "",
                        C21 = "",
                    };
                    datas.Add(datum);
                }
                string folderName = @"C:\RFID\Reports\";
                string filePath = "ReportAllNEW" + "_" + name + "_" + hod + ".xlsx";
                string folderPath = folderName + filePath;
                using (ExcelPackage excelPackage = new ExcelPackage())
                {
                    // Add a new worksheet to the workbook
                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");

                    // Add some content to the worksheet
                    for (int i = 0; i < datas.Count; i++)
                    {
                        for (int j = 0; j < 21; j++)
                        {
                            var propertyName = "C" + (j + 1);
                            var property = typeof(ReportsEPC).GetProperty(propertyName);
                            var propertyValue = property.GetValue(datas[i]);
                            worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Value = propertyValue;
                            worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                            worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Top.Color.SetColor(System.Drawing.Color.Black);
                            worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                            worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Bottom.Color.SetColor(System.Drawing.Color.Black);
                            worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                            worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Left.Color.SetColor(System.Drawing.Color.Black);
                            worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                            worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Right.Color.SetColor(System.Drawing.Color.Black);
                            worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                            // Set vertical alignment to center
                            worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            if (i == 0) // Nếu là hàng đầu tiên
                            {
                                worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                if (j == 0 || j == 3 || j == 5 || j == 6 || j == 7 || j == 8 || j == 9 || j == 10 || j == 11 || j == 12 || j == 14 || j == 15)
                                {
                                    worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 128, 0));
                                }
                                else if (j == 2 || j == 16 || j == 17)
                                {
                                    worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(0, 0, 255));
                                }
                                else if (j == 3 || j == 18 || j == 19 || j == 21)
                                {
                                    worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 255, 0));
                                }
                                else if (j == 13 || j == 20)
                                {
                                    worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 0, 0));
                                }
                                else
                                {
                                    worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 255, 255));
                                }


                            }
                        }
                    }
                    worksheet.View.FreezePanes(2, 1);
                    worksheet.View.FreezePanes(1, 6);
                    worksheet.Cells["A:AZ"].AutoFitColumns();
                    worksheet.Cells["A1:AZ1"].AutoFilter = true;
                    // Save the ExcelPackage to a file
                    FileInfo file = new FileInfo(folderPath);
                    excelPackage.SaveAs(file);
                }
                return Json(new { code = 200, }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception e)
            {
                return Json(new { code = 500, msg = "Sai !!!" + e.Message }, JsonRequestBehavior.AllowGet);
            }
        }
        [HttpPost]
        public JsonResult Save(string arr, string hodReports)
        {
            try
            {
                string Newarr = arr.Replace("Matched", "true").Replace("Mismatched", "true");
                var user = (User)Session["user"];
                var name = user.Name;
                var arrs = JsonConvert.DeserializeObject<EPCTOUPC.EpctoUpc[]>(Newarr);
                var datas = JsonConvert.DeserializeObject<Datum[]>(Newarr);
                // Lọc ra các bản ghi mới (có ID chưa tồn tại trong dữ liệu)
                var filteredRecords = datas.Where(r => !db.Data.Any(er => er.EPC == r.EPC)).ToList();
                filteredRecords.ForEach(record =>
                {
                    // Áp dụng hàm sửa đổi lên mỗi bản ghi ở đây
                    record.IdReports = hodReports;
                    record.So = db.Discrepancies.Where(x => x.UPC == record.UPC).ToList().LastOrDefault().So;
                    record.Po = db.Discrepancies.Where(x => x.UPC == record.UPC).ToList().LastOrDefault().Po;
                    record.Sku = db.Discrepancies.Where(x => x.UPC == record.UPC).ToList().LastOrDefault().Sku;
                    record.CreateDate = DateTime.Now;
                    record.CreateBy = name;
                });
                // Thêm các bản ghi mới vào danh sách hiện tại
                db.Data.AddRange(filteredRecords);

                HashSet<string> upcSet = new HashSet<string>();
                foreach (var obj in datas)
                {
                    string upc = (string)obj.UPC + "-" + obj.Status;
                    upcSet.Add(upc);
                }
                string[] uniqueUPCs = upcSet.ToArray();

                foreach (var item in uniqueUPCs)
                {
                    var u = item.Substring(0, item.IndexOf("-"));
                    var s = item.Substring(item.IndexOf("-") + 1);
                    var count = db.Discrepancies.Where(x => x.UPC == u && x.IdReports == hodReports).ToList();
                    if (count.Count() == 0)
                    {
                        continue;
                    }
                    int qts;
                    if (count.LastOrDefault().QtyScan == null)
                    {
                        qts = 0;
                    }
                    else
                    {
                        qts = int.Parse(count.LastOrDefault().QtyScan);
                    }
                    var sum = count.Sum(x => int.Parse(x.Qty));
                    var sumScan = int.Parse(count.LastOrDefault().QtyScan);
                    var q = arrs.OrderBy(x => x.UPC == u).ToList().LastOrDefault().Qty;
                    if (sum != int.Parse(q.Substring(0, q.IndexOf("/"))))
                    {
                        if (sumScan - int.Parse(q.Substring(0, q.IndexOf("/"))) <= 0)
                        {
                            foreach (var dis in count)
                            {
                                dis.QtyScan = q.Substring(0, q.IndexOf("/"));
                                db.SaveChanges();
                            }
                        }
                    }
                    else
                    {
                        foreach (var dis in count)
                        {
                            General general = new General()
                            {
                                So = dis.So,
                                Po = dis.Po,
                                Sku = dis.Sku,
                                IdReports = dis.IdReports,
                                CartonTo = dis.CartonTo,
                                UPC = dis.UPC,
                                Qty = dis.Qty,
                                CreateDate = DateTime.Now,
                                CreateBy = name
                            };
                            db.Generals.Add(general);
                            db.Discrepancies.Remove(dis);
                            db.SaveChanges();
                        }
                    }
                }
                return Json(new { code = 200, }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception e)
            {
                return Json(new { code = 500, msg = "Sai !!!" + e.Message }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpPost]
        public JsonResult Upload(HttpPostedFileBase file, string id)
        {
            try
            {
                if ((file != null) && (file.ContentLength > 0) && !string.IsNullOrEmpty(file.FileName))
                {
                    List<ReportsEPC> datas = new List<ReportsEPC>();
                    ReportsEPC dis = new ReportsEPC()
                    {
                        C1 = "CONSIGNEE",
                        C2 = "SHIPPER",
                        C3 = "SO NUMBER",
                        C4 = "PO NUMBER",
                        C5 = "SKU",
                        C6 = "CARTONID",
                        C7 = "UPC",
                        C8 = "TOTAL UNITS",
                        C9 = "SCANNING UNITS",
                        C10 = "RESULT",
                        C11 = "SCANNING DATE TIME",
                        C12 = "REMARK",
                        C13 = "HOD"
                    };
                    datas.Add(dis);
                    var reports = db.Reports.SingleOrDefault(x => x.Id == id);
                    var list = db.Discrepancies.Where(x => x.IdReports == id);

                    foreach (var item in list)
                    {
                        ReportsEPC datum = new ReportsEPC()
                        {
                            C1 = reports.Name,
                            C2 = reports.Shiper,
                            C3 = item.So,
                            C4 = item.Po,
                            C5 = item.Sku,
                            C6 = item.CartonTo,
                            C7 = item.UPC,
                            C8 = item.Qty,
                            C9 = item.QtyScan,
                            C10 = "Mismatched",
                            C11 = item.CreateDate == null ? "" : item.CreateDate.Value.Day + "-" + item.CreateDate.Value.Month + "-" + item.CreateDate.Value.Year + " " + item.CreateDate.Value.Hour + ":" + item.CreateDate.Value.Minute,
                            C12 = "",
                            C13 = reports.ModifyDate == null ? "" : reports.ModifyDate.Value.Day + "-" + reports.ModifyDate.Value.Month + "-" + reports.ModifyDate.Value.Year + " " + reports.ModifyDate.Value.Hour + ":" + reports.ModifyDate.Value.Minute,
                        };
                        datas.Add(datum);
                    }
                    string currentDirectory = HostingEnvironment.MapPath("~");
                    string fileName = file.FileName;
                    string fileContentType = file.ContentType;
                    byte[] fileBytes = new byte[file.ContentLength];
                    var data = file.InputStream.Read(fileBytes, 0, Convert.ToInt32(file.ContentLength));
                    string[] upcDiscrepancy = new string[0];
                    for (int i = 1; i < datas.Count; i++)
                    {
                        int index = Array.IndexOf(upcDiscrepancy, Array.Find(upcDiscrepancy, element => element.Contains(datas[i].C7)));
                        if (index >= 0)
                        {
                            var sl = upcDiscrepancy[index].Substring(upcDiscrepancy[index].IndexOf("-") + 1);
                            upcDiscrepancy[index] = datas[i].C7 + "-" + (int.Parse(datas[i].C8) + int.Parse(sl));
                        }
                        else
                        {
                            Array.Resize(ref upcDiscrepancy, upcDiscrepancy.Length + 1);
                            upcDiscrepancy[upcDiscrepancy.Length - 1] = datas[i].C7 + "-" + datas[i].C8;
                        }
                    }
                    if (file.ContentType.Contains("csv"))
                    {
                        var count = 0;
                        string csvFilePath = @"C:\RFID\FileCSVhandheld\" + file.FileName + "";
                        using (var reader = new StreamReader(csvFilePath))
                        {
                            List<Gerenal.gerenal> generals = new List<Gerenal.gerenal>();
                            List<EPCTOUPC.EpctoUpc> epctoUpcs = new List<EPCTOUPC.EpctoUpc>();
                            while (!reader.EndOfStream)
                            {
                                count++;
                                string line = reader.ReadLine();
                                string[] fields = line.Split(',');
                                // Xử lý các trường trong dòng
                                if (count >= 7)
                                {

                                    try
                                    {
                                        var epc = fields[0];

                                        if (epc == null || epc.Length <= 21)
                                        {
                                            continue;
                                        }
                                        EPCTOUPC.EpctoUpc epctoUpc = new EPCTOUPC.EpctoUpc()
                                        {
                                            EPC = epc,
                                            UPC = epctoupc(epc)
                                        };
                                        epctoUpcs.Add(epctoUpc);
                                        if (generals.Any(item => item.Upc == epctoupc(epc)))
                                        {
                                            var general = generals.SingleOrDefault(x => x.Upc == epctoupc(epc));
                                            int qty = int.Parse(general.Qty);
                                            qty++; general.Qty = qty.ToString();
                                        }
                                        else
                                        {
                                            Gerenal.gerenal gerenal = new Gerenal.gerenal()
                                            {
                                                Upc = epctoupc(epc),
                                                Qty = "1",
                                                Status = true
                                            };
                                            generals.Add(gerenal);
                                        }
                                    }
                                    catch (DbEntityValidationException ex)
                                    {
                                        foreach (var error in ex.EntityValidationErrors)
                                        {
                                            foreach (var validationError in error.ValidationErrors)
                                            {
                                                Console.WriteLine("Lỗi xác thực: {0}", validationError.ErrorMessage);
                                            }
                                        }
                                    }
                                    foreach (var item in upcDiscrepancy)
                                    {
                                        var upc = item.Substring(0, item.IndexOf("-"));
                                        var qty = item.Substring(item.IndexOf("-") + 1);
                                        if (generals.Any(x => x.Upc == upc))
                                        {
                                            if (int.Parse(generals.SingleOrDefault(x => x.Upc == upc).Qty) == int.Parse(qty))
                                            {
                                                foreach (var status in epctoUpcs)
                                                {
                                                    if (status.UPC == upc)
                                                    {
                                                        status.Status = "Matched";
                                                        status.Qty = generals.SingleOrDefault(x => x.Upc == upc).Qty + "/" + qty;

                                                    }
                                                }
                                            }
                                            else if (int.Parse(generals.SingleOrDefault(x => x.Upc == upc).Qty) > int.Parse(qty))
                                            {
                                                foreach (var status in epctoUpcs)
                                                {
                                                    if (status.UPC == upc)
                                                    {
                                                        status.Status = "Mismatched";
                                                        status.Qty = generals.SingleOrDefault(x => x.Upc == upc).Qty + "/" + qty;

                                                    }
                                                }
                                            }
                                            else
                                            {
                                                foreach (var status in epctoUpcs)
                                                {
                                                    if (status.UPC == upc)
                                                    {
                                                        status.Status = "Mismatched";
                                                        status.Qty = generals.SingleOrDefault(x => x.Upc == upc).Qty + "/" + qty;


                                                    }
                                                }
                                            }
                                        }
                                        else
                                        {
                                            foreach (var status in epctoUpcs)
                                            {
                                                if (status.UPC == upc)
                                                {
                                                    status.Status = "error";
                                                    status.Qty = generals.SingleOrDefault(x => x.Upc == upc).Qty + "/" + qty;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            return Json(new { code = 200, msg = epctoUpcs.ToList() }, JsonRequestBehavior.AllowGet);
                        }
                    }
                    else
                    {

                    }
                    using (var package = new ExcelPackage(file.InputStream))
                    {
                        ExcelWorksheet currentSheet = package.Workbook.Worksheets.First();
                        var workSheet = currentSheet;
                        var noOfCol = workSheet.Dimension.End.Column;
                        var noOfRow = workSheet.Dimension.End.Row;
                        List<Gerenal.gerenal> generals = new List<Gerenal.gerenal>();
                        List<EPCTOUPC.EpctoUpc> epctoUpcs = new List<EPCTOUPC.EpctoUpc>();
                        for (int rowIterator = 7; rowIterator <= noOfRow; rowIterator++)
                        {
                            try
                            {
                                var epc = workSheet.Cells[rowIterator, 1].Value == null ? null : workSheet.Cells[rowIterator, 1].Value.ToString();
                                if (epc == null || epc.Length <= 21)
                                {
                                    continue;
                                }
                                EPCTOUPC.EpctoUpc epctoUpc = new EPCTOUPC.EpctoUpc()
                                {
                                    EPC = epc,
                                    UPC = epctoupc(epc)
                                };
                                epctoUpcs.Add(epctoUpc);
                                if (generals.Any(item => item.Upc == epctoupc(epc)))
                                {
                                    var general = generals.SingleOrDefault(x => x.Upc == epctoupc(epc));
                                    int qty = int.Parse(general.Qty);
                                    qty++;
                                    general.Qty = qty.ToString();
                                }
                                else
                                {
                                    Gerenal.gerenal gerenal = new Gerenal.gerenal()
                                    {
                                        Upc = epctoupc(epc),
                                        Qty = "1",
                                        Status = true
                                    };
                                    generals.Add(gerenal);
                                }
                            }
                            catch (DbEntityValidationException ex)
                            {
                                foreach (var error in ex.EntityValidationErrors)
                                {
                                    foreach (var validationError in error.ValidationErrors)
                                    {
                                        Console.WriteLine("Lỗi xác thực: {0}", validationError.ErrorMessage);
                                    }
                                }
                            }
                        }
                        foreach (var item in upcDiscrepancy)
                        {
                            var upc = item.Substring(0, item.IndexOf("-"));
                            var qty = item.Substring(item.IndexOf("-") + 1);
                            var Info = db.Discrepancies.OrderBy(x => x.UPC == upc && x.IdReports == id).ToList().LastOrDefault();
                            if (generals.Any(x => x.Upc == upc))
                            {
                                if (int.Parse(generals.SingleOrDefault(x => x.Upc == upc).Qty) == int.Parse(qty))
                                {
                                    foreach (var status in epctoUpcs)
                                    {
                                        if (status.UPC == upc)
                                        {
                                            status.Status = "Matched";
                                            status.Qty = generals.SingleOrDefault(x => x.Upc == upc).Qty + "/" + qty;

                                        }
                                    }
                                }
                                else if (int.Parse(generals.SingleOrDefault(x => x.Upc == upc).Qty) > int.Parse(qty))
                                {
                                    foreach (var status in epctoUpcs)
                                    {
                                        if (status.UPC == upc)
                                        {
                                            status.Status = "Mismatched";
                                            status.Qty = generals.SingleOrDefault(x => x.Upc == upc).Qty + "/" + qty;

                                        }
                                    }
                                }
                                else
                                {
                                    foreach (var status in epctoUpcs)
                                    {
                                        if (status.UPC == upc)
                                        {
                                            status.Status = "Mismatched";
                                            status.Qty = generals.SingleOrDefault(x => x.Upc == upc).Qty + "/" + qty;

                                        }
                                    }
                                }
                            }
                            else
                            {
                                foreach (var status in epctoUpcs)
                                {
                                    if (status.UPC == upc)
                                    {
                                        status.Status = "error";
                                        status.Qty = generals.SingleOrDefault(x => x.Upc == upc).Qty + "/" + qty;
                                    }
                                }
                            }
                        }
                        return Json(new { code = 200, msg = epctoUpcs.ToList() }, JsonRequestBehavior.AllowGet);
                    }
                }
                return Json(new { code = 300, }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception e)
            {
                return Json(new { code = 500, msg = rm.GetString("false").ToString() + " !!!" + e.Message }, JsonRequestBehavior.AllowGet);
            }
        }

        public static string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = string.Empty;
            int modulo;
            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }

        public static int sumDate(string date)
        {
            var y = date == "" ? "" : date.Substring(0, date.IndexOf("-"));
            var m = date == "" ? "" : date.Substring(date.IndexOf("-") + 1, 2);
            var d = date == "" ? "" : date.Substring(date.LastIndexOf("-") + 1, 2);
            var s = date == "" ? 0 : int.Parse(d) + (int.Parse(m) * 30) + int.Parse(y);
            return s;
        }

        public static int Quy(int month)
        {
            return (month + 2) / 3;
        }

        public void Report(string[] data, string name, string hod, string type)
        {
            var datas = JsonConvert.DeserializeObject<string[][]>(data[0]);
            string folderName = @"C:\RFID\Reports\";
            string filePath = type + "_" + name + "_" + hod + ".xlsx";
            string folderPath = folderName + filePath;
            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                // Add a new worksheet to the workbook
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");

                // Add some content to the worksheet
                for (int i = 0; i < datas.Length; i++)
                {
                    for (int j = 0; j < datas[0].Length; j++)
                    {
                        worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Value = datas[i][j];
                        worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Top.Style = ExcelBorderStyle.Hair;
                        worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Top.Color.SetColor(System.Drawing.Color.Black);
                        worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                        worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Bottom.Color.SetColor(System.Drawing.Color.Black);
                        worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Left.Style = ExcelBorderStyle.Hair;
                        worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Left.Color.SetColor(System.Drawing.Color.Black);
                        worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Right.Style = ExcelBorderStyle.Hair;
                        worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Border.Right.Color.SetColor(System.Drawing.Color.Black);
                        if (i == 0) // Nếu là hàng đầu tiên
                        {
                            worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 255, 0));
                        }
                    }
                }
                worksheet.View.FreezePanes(2, 1);
                // Save the ExcelPackage to a file
                FileInfo file = new FileInfo(folderPath);
                excelPackage.SaveAs(file);
            }
        }
        public Gerenal.gerenal hasUpc(string upc, string hodReports)
        {
            Gerenal.gerenal rsl = new Gerenal.gerenal();
            if (db.Generals.OrderBy(x => x.UPC == upc && x.IdReports == hodReports).ToList().LastOrDefault() == null)
            {
                var arr = db.Discrepancies.OrderBy(x => x.UPC == upc && x.IdReports == hodReports).ToList().LastOrDefault();
                rsl.So = arr.So;
                rsl.Po = arr.Po;
                rsl.Sku = arr.Sku;
                return rsl;
            }
            else
            {
                var arr = db.Generals.OrderBy(x => x.UPC == upc && x.IdReports == hodReports).ToList().LastOrDefault();
                rsl.So = arr.So;
                rsl.Po = arr.Po;
                rsl.Sku = arr.Sku;
                return rsl;
            }
            return rsl;
        }
        public totalQTY TotalQTYReports(string id)
        {
            totalQTY totalQTY = new totalQTY();
            List<string> code = new List<string>();
            var general = db.Generals.Where(x => x.IdReports == id).ToList();
            var discrepancies = db.Discrepancies.Where(x => x.IdReports == id).ToList();
            totalQTY.QtyCtn = db.DataScanPhysicals.Where(x => x.IdReports == id).Count();
            foreach (var item in general)
            {
                totalQTY.Qty += int.Parse(item.Qty);
                totalQTY.QtyScan += int.Parse(item.Qty);
            }
            foreach (var item in discrepancies)
            {
                totalQTY.Qty += int.Parse(item.Qty);
                if (item.QtyScan != null)
                {
                    if (!code.Contains(item.Sku))
                    {
                        totalQTY.QtyScan += int.Parse(item.QtyScan);
                        code.Add(item.Sku);
                    }

                }
            }
            return totalQTY;
        }

        public string checkPort(string id)
        {
            var check = db.EPCDiscrepancies.FirstOrDefault(x => x.IdReports == id);
            var general = db.Generals.FirstOrDefault(x => x.IdReports == id);
            var discrepancies = db.Discrepancies.FirstOrDefault(x => x.IdReports == id);
            if (check == null)
            {
                if (general == null)
                {
                    return discrepancies.port;
                }
                else
                {
                    return general.port;
                }
            }
            else
            {
                return check.port;
            }
        }
        public string epctoupc(string epc)
        {
            //var GetBitEnd = int.Parse((string)Session["GetBitEnd"]);
            //var GetBitSGTIN = int.Parse((string)Session["GetBitSGTIN"]);
            //var GetBitItemRef = int.Parse((string)Session["GetBitItemRef"]);
            //var AddCharacter = int.Parse((string)Session["AddCharacter"]);
            //var AddCharacterstring = (string)Session["AddCharacterstring"];
            //var TakeAbsoluteValue = int.Parse((string)Session["TakeAbsoluteValue"]);
            //var CheckDigits = int.Parse((string)Session["CheckDigit"]);
            //string EPC = epc;
            //string EPCIN = "";
            //string SGTIN = "";
            //string ItemRef = "";
            //string Result = "";
            //string UPC = "";
            //int i, SGTINResult, ItemRefResult, CheckDigit = 0;
            //for (i = 0; i < EPC.Length; i++)
            //{
            //    EPCIN += Convert.ToString(Convert.ToInt32(EPC.Substring(i, 1), 16), 2).PadLeft(4, '0');
            //}
            //EPCIN = EPCIN.Substring(EPCIN.Length - GetBitEnd);
            //SGTIN = EPCIN.Substring(0, GetBitSGTIN);
            //ItemRef = EPCIN.Substring(GetBitSGTIN, GetBitItemRef);
            //SGTINResult = 0;
            //for (i = 1; i < SGTIN.Length; i++)
            //{
            //    SGTINResult += Convert.ToInt32(SGTIN.Substring(i, 1)) * (int)Math.Pow(2, SGTIN.Length - i - 1);
            //}
            //ItemRefResult = 0;
            //for (i = 1; i < ItemRef.Length; i++)
            //{
            //    ItemRefResult += Convert.ToInt32(ItemRef.Substring(i, 1)) * (int)Math.Pow(2, ItemRef.Length - i - 1);
            //}
            //Result = SGTINResult.ToString() + (AddCharacterstring + ItemRefResult).Substring(Math.Max(0, (AddCharacterstring + ItemRefResult).Length - AddCharacter));
            //CheckDigit = 0;
            //for (i = 1; i <= TakeAbsoluteValue; i++)
            //{
            //    if (Result.Length > Math.Abs(i - TakeAbsoluteValue))
            //    {
            //        if (i % 2 != 0)
            //        {
            //            CheckDigit += 3 * Convert.ToInt32(Result.Substring(Result.Length - Math.Abs(i - TakeAbsoluteValue) - 1, 1));
            //        }
            //        else
            //        {
            //            CheckDigit += Convert.ToInt32(Result.Substring(Result.Length - Math.Abs(i - TakeAbsoluteValue) - 1, 1));
            //        }
            //    }
            //}
            //CheckDigit = Convert.ToInt32(Math.Ceiling((double)CheckDigit / CheckDigits) * CheckDigits) - CheckDigit; UPC = Result + CheckDigit; return UPC;
            string EPC = epc;
            string EPCIN = "";
            string SGTIN = "";
            string ItemRef = "";
            string Result = "";
            string UPC = "";
            int i, SGTINResult, ItemRefResult, CheckDigit = 0;

            for (i = 0; i < EPC.Length; i++)
            {
                EPCIN += Convert.ToString(Convert.ToInt32(EPC.Substring(i, 1), 16), 2).PadLeft(4, '0');
            }

            EPCIN = EPCIN.Substring(EPCIN.Length - 82);
            SGTIN = EPCIN.Substring(0, 24);
            ItemRef = EPCIN.Substring(24, 20);

            SGTINResult = 0;
            for (i = 1; i < SGTIN.Length; i++)
            {
                SGTINResult += Convert.ToInt32(SGTIN.Substring(i, 1)) * (int)Math.Pow(2, SGTIN.Length - i - 1);
            }

            ItemRefResult = 0;
            for (i = 1; i < ItemRef.Length; i++)
            {
                ItemRefResult += Convert.ToInt32(ItemRef.Substring(i, 1)) * (int)Math.Pow(2, ItemRef.Length - i - 1);
            }

            Result = SGTINResult.ToString() + ("00000" + ItemRefResult).Substring(Math.Max(0, ("00000" + ItemRefResult).Length - 5));

            CheckDigit = 0;
            for (i = 1; i <= 17; i++)
            {
                if (Result.Length > Math.Abs(i - 17))
                {
                    if (i % 2 != 0)
                    {
                        CheckDigit += 3 * Convert.ToInt32(Result.Substring(Result.Length - Math.Abs(i - 17) - 1, 1));
                    }
                    else
                    {
                        CheckDigit += Convert.ToInt32(Result.Substring(Result.Length - Math.Abs(i - 17) - 1, 1));
                    }
                }
            }

            CheckDigit = Convert.ToInt32(Math.Ceiling((double)CheckDigit / 10) * 10) - CheckDigit;
            UPC = Result + CheckDigit;

            return UPC;
        }

    }
}