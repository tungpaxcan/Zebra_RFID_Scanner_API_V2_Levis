using Microsoft.Ajax.Utilities;
using Microsoft.AspNet.SignalR;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Diagnostics;
using System.IO.Ports;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Resources;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using Zebra_RFID_Scanner.Models;

namespace Zebra_RFID_Scanner.Controllers
{
    public class HomeController : BaseController
    {
        private Entities db = new Entities();
        ResourceManager rm = new ResourceManager("Zebra_RFID_Scanner.Resources.Resource", typeof(Resources.Resource).Assembly);
        public System.Web.Mvc.ActionResult Index()
        {

            return View();
        }
        public System.Web.Mvc.ActionResult Index2()
        {

            return View();
        }
        public System.Web.Mvc.ActionResult Help()
        {

            return View();
        }

        public System.Web.Mvc.ActionResult Advanced()
        {

            return View();
        }
        [System.Web.Mvc.HttpGet]
        public JsonResult xmlValidate(string so)
        {
            try
            {
                var validate = (from b in db.XMLValidates.Where(x => x.So == so)
                            select new
                            {
                                id = b.Id,
                                ctn = b.Ctn
                            }).ToList();
                return Json(new { code = 200, validate = validate }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception e)
            {
                return Json(new { code = 500, msg = rm.GetString("false") + " " + e.Message }, JsonRequestBehavior.AllowGet);
            }
        }
        [System.Web.Mvc.HttpGet]
        public JsonResult UserAll()
        {
            try
            {
                var user = (from b in db.Users.Where(x => x.Id > 0)
                            select new
                            {
                                id = b.Id,
                                name = b.Name,
                            }).ToList();
                return Json(new { code = 200,user=user}, JsonRequestBehavior.AllowGet);
            }
            catch (Exception e)
            {
                return Json(new { code = 500, msg = rm.GetString("false") + " " + e.Message }, JsonRequestBehavior.AllowGet);
            }
        }
        [System.Web.Mvc.HttpGet]
        public JsonResult UserSession()
        {
            try
            {
                var datetime = DateTime.Now.Day + (DateTime.Now.Month * 30) + DateTime.Now.Year;//2237
                var re = db.Reports.Where(x => (datetime)-(x.CreateDate.Value.Day + (x.CreateDate.Value.Month * 30) + x.CreateDate.Value.Year) >= 56).Count();
                for (int i = 0; i < re; i++)
                {
                    var idRe = db.Reports.OrderBy(x => (datetime)-(x.CreateDate.Value.Day + (x.CreateDate.Value.Month * 30) + x.CreateDate.Value.Year) >= 56).ToList().LastOrDefault().Id;
                    Dele.DeleteReports(idRe);
                }
                //if (datetime>=2237)
                //{
                //    var user = db.Users.ToList();
                //    foreach(var u in user)
                //    {
                //        u.IdFX = null;
                //        db.SaveChanges();
                //    }
                //}
                var c = (User)Session["user"];
                Session["GetBitEnd"] = c.GetBitEnd;
                Session["GetBitSGTIN"] = c.GetBitGTIN;
                Session["GetBitItemRef"] = c.GetBitItemRef;
                Session["GetBitItemRefcount"] = int.Parse(c.GetBitItemRef)+int.Parse(c.GetBitGTIN);
                Session["AddCharacter"] = c.AddCharacter;
                string str = "";
                for(int i = 0; i < int.Parse(c.AddCharacter); i++)
                {
                    str += "0";
                }
                Session["AddCharacterstring"] = str;
                Session["TakeAbsoluteValue"] = c.TakeAbsoluteValue;
                Session["CheckDigit"] = c.CheckDigit;
                return Json(new { code = 200, name = c.Name,id=c.Id,admin=c.Admin}, JsonRequestBehavior.AllowGet);
            }
            catch (Exception e)
            {
                return Json(new { code = 500, msg = rm.GetString("false")+" " + e.Message }, JsonRequestBehavior.AllowGet);
            }
        }
        [System.Web.Mvc.HttpPost]
        public JsonResult EditFormula(User fromData)
        {
            try
            {
                var user = (User)Session["user"];
                var us = db.Users.Find(user.Id);
                us.GetBitEnd = fromData.GetBitEnd;
                us.GetBitGTIN = fromData.GetBitGTIN;
                us.GetBitItemRef = fromData.GetBitItemRef;
                us.AddCharacter = fromData.AddCharacter;
                us.TakeAbsoluteValue = fromData.TakeAbsoluteValue;
                us.CheckDigit = fromData.CheckDigit;
                db.SaveChanges();
                Session["user"] = null;
                return Json(new { code = 200,}, JsonRequestBehavior.AllowGet);
            }
            catch (Exception e)
            {
                return Json(new { code = 500, msg = rm.GetString("false") + " " + e.Message }, JsonRequestBehavior.AllowGet);
            }
        }
        [System.Web.Mvc.HttpPost]
        public JsonResult ResetDefault()
        {
            try
            {
                var user = (User)Session["user"];
                var us = db.Users.Find(user.Id);
                us.GetBitEnd = "82";
                us.GetBitGTIN = "24";
                us.GetBitItemRef = "20";
                us.AddCharacter = "5";
                us.TakeAbsoluteValue = "17";
                us.CheckDigit = "10";
                db.SaveChanges();
                Session["user"] = null;
                return Json(new { code = 200, }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception e)
            {
                return Json(new { code = 500, msg = rm.GetString("false") + " " + e.Message }, JsonRequestBehavior.AllowGet);
            }
        }
        [System.Web.Mvc.HttpGet]
        public JsonResult CheckPreASN(string name,bool type)
        {
            try
            {
                if (db.Reports.Find(name) != null)
                {
                    return Json(new { code = "error", msg = "already in the system, can't save" }, JsonRequestBehavior.AllowGet);
                }
                else
                {
                    if(type == false)
                    {
                        return Json(new { code = 200,url = "/Home/Index" }, JsonRequestBehavior.AllowGet);
                    }
                    else
                    {
                        return Json(new { code = 200, url = "/Home/Index2" }, JsonRequestBehavior.AllowGet);
                    }
                  
                }
            }
            catch (Exception e)
            {
                return Json(new { code = 500, msg = rm.GetString("false") + " " + e.Message }, JsonRequestBehavior.AllowGet);
            }
        }
        [System.Web.Mvc.HttpPost]
        public JsonResult Save(string idReports, string Ctn, string ctnError, string epcToUpc, string Po, string So, string Sku, string info, string Consignee, string Shipper,string TimeStart,string EPCDiscrepancy)
        {
            try
            {
                var time = DateTime.Now;
                var nameUser = "";
                var fxConnect = "";
                if (TimeStart == "")
                {
                    time = DateTime.Now;
                }
                else
                {
                    time = DateTime.Parse(TimeStart);
                }
                if (Session?["user"] != null)
                {
                    var user = (User)Session["user"];
                    nameUser = user.Name;
                    fxConnect = user.FXconnect?.Name;
                }
               
                Ctn = Ctn.Replace("Carton to", "CartonTo");
                Ctn = Ctn.Replace("UPC QTY", "UPCQty");
                ctnError = ctnError.Replace("Carton to", "CartonTo");
                ctnError = ctnError.Replace("UPC QTY", "UPCQty");
                var ctnErrors = string.IsNullOrEmpty(ctnError) ? new DataScanPhysical[] { } : JsonConvert.DeserializeObject<DataScanPhysical[]>(ctnError);
                var Ctns = string.IsNullOrEmpty(Ctn) ? new DataScanPhysical[] { } : JsonConvert.DeserializeObject<DataScanPhysical[]>(Ctn);
                var generals = string.IsNullOrEmpty(Ctn) ? new General[] { } : JsonConvert.DeserializeObject<General[]>(Ctn);
                var discrepancy = string.IsNullOrEmpty(Ctn) ? new Discrepancy[] { } : JsonConvert.DeserializeObject<Discrepancy[]>(Ctn);
                var epcToUpcs = string.IsNullOrEmpty(epcToUpc) ? new Datum[] { } : JsonConvert.DeserializeObject<Datum[]>(epcToUpc);
                var EPCDiscrepancys = string.IsNullOrEmpty(EPCDiscrepancy) ? new EPCDiscrepancy[] { }: JsonConvert.DeserializeObject<EPCDiscrepancy[]>(EPCDiscrepancy);
                var infos = JsonConvert.DeserializeObject<Gerenal.gerenal[]>(info);
                if (db.Reports.Find(idReports) != null) { return Json(new { code = "error", msg = "already in the system, can't save" }, JsonRequestBehavior.AllowGet); }
                Report report = new Report();
                report.Id = idReports;
                report.Name = Consignee;
                report.Shiper = Shipper;
                report.So = So;
                report.Po = Po;
                report.Sku = Sku;
                report.CreateDate = DateTime.Now;
                report.CreateBy = nameUser;
                report.Status = false;
                report.TimeStart = time;
                db.Reports.Add(report);
                epcToUpcs.ForEach(record =>
                {
                    // Áp dụng hàm sửa đổi lên mỗi bản ghi ở đây
                    record.CreateDate = DateTime.Now;
                    record.CreateBy = nameUser;
                    if (string.IsNullOrEmpty(record.deviceNum))
                    {
                        record.deviceNum = fxConnect; // Sử dụng ?. để tránh lỗi null reference nếu user.FXconnect là null
                    }
                    record.Status = true;
                });
                db.Data.AddRange(epcToUpcs);
                EPCDiscrepancys.ForEach(record =>
                {
                    // Áp dụng hàm sửa đổi lên mỗi bản ghi ở đây
                    record.CreateDate = DateTime.Now;
                    record.CreateBy = nameUser;
                    record.IdReports = idReports;
                    record.Status = true;
                    if (string.IsNullOrEmpty(record.deviceNum))
                    {
                        record.deviceNum = fxConnect; // Sử dụng ?. để tránh lỗi null reference nếu user.FXconnect là null
                    }
                });
                db.EPCDiscrepancies.AddRange(EPCDiscrepancys);
                Ctns.ForEach(record =>
                {
                    // Áp dụng hàm sửa đổi lên mỗi bản ghi ở đây
                    record.IdReports = idReports;
                    record.Status = true;
                    record.CreateDate = DateTime.Now;
                    record.CreateBy = nameUser;
                });
                db.DataScanPhysicals.AddRange(Ctns);
                ctnErrors.ForEach(record =>
                {
                    // Áp dụng hàm sửa đổi lên mỗi bản ghi ở đây
                    record.IdReports = idReports;
                    record.Status = false;
                    record.CreateDate = DateTime.Now;
                    record.CreateBy = nameUser;
                });
                db.DataScanPhysicals.AddRange(ctnErrors);

                generals.ForEach(record =>
                {
                    // Áp dụng hàm sửa đổi lên mỗi bản ghi ở đây
                    record.IdReports = idReports;
                    record.CreateDate = DateTime.Now;
                    record.CreateBy = nameUser;
                    if (string.IsNullOrEmpty(record.deviceNum))
                    {
                        record.deviceNum = fxConnect; // Sử dụng ?. để tránh lỗi null reference nếu user.FXconnect là null
                    }
                });
                db.Generals.AddRange(generals.Where(item => item.Status == true));
                discrepancy.ForEach(record =>
                {
                    // Áp dụng hàm sửa đổi lên mỗi bản ghi ở đây
                    if (record.QtyScan == null)
                    {
                        record.QtyScan = "0";
                    }
                    record.IdReports = idReports;
                    record.CreateDate = DateTime.Now;
                    record.CreateBy = nameUser;
                    if (string.IsNullOrEmpty(record.deviceNum))
                    {
                        record.deviceNum = fxConnect; // Sử dụng ?. để tránh lỗi null reference nếu user.FXconnect là null
                    }
                });
                db.Discrepancies.AddRange(discrepancy.Where(item => item.Status == false));

                db.SaveChanges();
                return Json(new { code = 200, msg = rm.GetString("success").ToString(), }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception e)
            {
                return Json(new { code = 500, msg = "Error " + e.Message }, JsonRequestBehavior.AllowGet);
            }
        }


        [System.Web.Mvc.HttpPost]
        public JsonResult Edit(string idReports, string Ctn, string epcToUpc, string Po, string So, string Sku, string general, string info)
        {
            try
            {
                var user = (User)Session["user"];
                var nameUser = user.Name;
                Ctn = Ctn.Replace("Carton to", "CartonTo");
                Ctn = Ctn.Replace("UPC QTY", "UPCQty");
                var Ctns = JsonConvert.DeserializeObject<Ctn.Carton[]>(Ctn);
                var epcToUpcs = JsonConvert.DeserializeObject<EPCTOUPC.EpctoUpc[]>(epcToUpc);
                var generals = JsonConvert.DeserializeObject<Gerenal.gerenal[]>(general);
                var infos = JsonConvert.DeserializeObject<Gerenal.gerenal[]>(info);
                for (int i = 0; i < Ctns.Length; i++)
                {
                    var c = Ctns[i].CartonTo;
                    var dataScanPhysical = db.DataScanPhysicals.SingleOrDefault(x => x.IdReports == idReports && x.CartonTo == c);
                    dataScanPhysical.CreateDate = DateTime.Now;
                    dataScanPhysical.Status = true;
                    db.SaveChanges();
                }
                for (int i = 0; i < epcToUpcs.Length; i++)
                {
                    Zebra_RFID_Scanner.Models.Datum datum = new Datum();
                    datum.IdReports = idReports;
                    datum.So = getInfo(infos, epcToUpcs[i].UPC).So;
                    datum.Po = getInfo(infos, epcToUpcs[i].UPC).Po;
                    datum.Sku = getInfo(infos, epcToUpcs[i].UPC).Sku;
                    datum.EPC = epcToUpcs[i].EPC;
                    datum.UPC = epcToUpcs[i].UPC;
                    datum.CreateDate = DateTime.Now;
                    datum.CreateBy = nameUser;
                    datum.Status = true;
                    db.Data.Add(datum);
                    db.SaveChanges();
                }
                for (int i = 0; i < Ctns.Length; i++)
                {
                    if (Ctns[i].status == true)
                    {
                        Zebra_RFID_Scanner.Models.General general1 = new Zebra_RFID_Scanner.Models.General();
                        general1.IdReports = idReports;
                        general1.CartonTo = Ctns[i].CartonTo;
                        general1.So = getInfo(infos, Ctns[i].UPC).So;
                        general1.Po = getInfo(infos, Ctns[i].UPC).Po;
                        general1.Sku = getInfo(infos, Ctns[i].UPC).Sku;
                        general1.UPC = Ctns[i].UPC;
                        general1.Qty = Ctns[i].UPCQty;
                        general1.CreateDate = DateTime.Now;
                        general1.CreateBy = nameUser;
                        general1.Status = true;
                        db.Generals.Add(general1);
                        db.SaveChanges();
                    }
                    else
                    {
                        Zebra_RFID_Scanner.Models.Discrepancy discrepancy = new Zebra_RFID_Scanner.Models.Discrepancy();
                        discrepancy.IdReports = idReports;
                        discrepancy.CartonTo = Ctns[i].CartonTo;
                        discrepancy.So = getInfo(infos, Ctns[i].UPC).So;
                        discrepancy.Po = getInfo(infos, Ctns[i].UPC).Po;
                        discrepancy.Sku = getInfo(infos, Ctns[i].UPC).Sku;
                        discrepancy.UPC = Ctns[i].UPC;
                        discrepancy.Qty = Ctns[i].UPCQty;
                        discrepancy.QtyScan = Ctns[i].QtyScan;
                        discrepancy.CreateDate = DateTime.Now;
                        discrepancy.CreateBy = nameUser;
                        discrepancy.Status = true;
                        db.Discrepancies.Add(discrepancy);
                        db.SaveChanges();
                    }
                }
                return Json(new { code = 200, msg = rm.GetString("success").ToString() }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception e)
            {
                return Json(new { code = 500, msg = rm.GetString("false").ToString() + " " + e.Message }, JsonRequestBehavior.AllowGet);
            }
        }
        [System.Web.Mvc.HttpPost]
        public JsonResult EditAll(string idReports, string Ctn, string epcToUpc, string Po, string So, string Sku, string info)
        {
            try
            {
                var user = (User)Session["user"];
                var nameUser = user.Name;
                Ctn = Ctn.Replace("Carton to", "CartonTo");
                Ctn = Ctn.Replace("UPC QTY", "UPCQty");
                var Ctns = JsonConvert.DeserializeObject<DataScanPhysical[]>(Ctn);
                var epcToUpcs = JsonConvert.DeserializeObject<Datum[]>(epcToUpc);
                var generals = JsonConvert.DeserializeObject<General[]>(Ctn);
                var discrepancy = JsonConvert.DeserializeObject<Discrepancy[]>(Ctn);
                var infos = JsonConvert.DeserializeObject<Gerenal.gerenal[]>(info);
                Dele.DeleteDiscrepancies(idReports);
                Dele.DeleteGenerals(idReports);
                foreach(var item in Ctns)
                {
                    var carton = db.DataScanPhysicals.FirstOrDefault(x => x.CartonTo == item.CartonTo && x.IdReports == idReports && x.Status == false && x.Po == item.Po);
                    if(carton != null)
                    {
                        carton.Status = true;
                    }
                }
                epcToUpcs.ForEach(record =>
                {
                    // Áp dụng hàm sửa đổi lên mỗi bản ghi ở đây
                    record.CreateDate = DateTime.Now;
                    record.CreateBy = nameUser;
                    record.Status = true;
                    if (string.IsNullOrEmpty(record.deviceNum))
                    {
                        record.deviceNum = user.FXconnect?.Name; // Sử dụng ?. để tránh lỗi null reference nếu user.FXconnect là null
                    }
                });
                db.Data.AddRange(epcToUpcs);

                generals.ForEach(record =>
                {
                    // Áp dụng hàm sửa đổi lên mỗi bản ghi ở đây
                    record.IdReports = idReports;
                    record.CreateDate = DateTime.Now;
                    record.CreateBy = nameUser;
                    if (string.IsNullOrEmpty(record.deviceNum))
                    {
                        record.deviceNum = user.FXconnect?.Name; // Sử dụng ?. để tránh lỗi null reference nếu user.FXconnect là null
                    }
                });
                db.Generals.AddRange(generals.Where(item => item.Status == true));
                discrepancy.ForEach(record =>
                {
                    // Áp dụng hàm sửa đổi lên mỗi bản ghi ở đây
                    record.IdReports = idReports;
                    record.CreateDate = DateTime.Now;
                    record.CreateBy = nameUser;
                    if (string.IsNullOrEmpty(record.deviceNum))
                    {
                        record.deviceNum = user.FXconnect?.Name; // Sử dụng ?. để tránh lỗi null reference nếu user.FXconnect là null
                    }
                });
                db.Discrepancies.AddRange(discrepancy.Where(item => item.Status == false));
                foreach(var item in epcToUpcs)
                {
                    var epcDiscrepancy = db.EPCDiscrepancies.SingleOrDefault(x => x.Id == item.EPC && idReports == item.IdReports);
                    if(epcDiscrepancy != null)
                    {
                        db.EPCDiscrepancies.Remove(epcDiscrepancy);
                    }
                }
                db.SaveChanges();
                return Json(new { code = 200, msg = rm.GetString("success").ToString() }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception e)
            {
                return Json(new { code = 500, msg = rm.GetString("false").ToString() + " " + e.Message }, JsonRequestBehavior.AllowGet);
            }
        }
        public static Gerenal.gerenal getInfo(Gerenal.gerenal[] general,string upc)
        {
            Gerenal.gerenal general1 = new Gerenal.gerenal();
            for(int i = 0; i < general.Length; i++)
            {
                if (general[i].Upc.Trim() == upc.Trim())
                {
                    general1.So = general[i].So;                    
                    general1.Po = general[i].Po;                    
                    general1.Sku = general[i].Sku;                    
                }
            }
            return general1;
        }
        [System.Web.Mvc.HttpGet]
        public JsonResult Page_Load()
        { // Kiểm tra tính còn hiệu lực của phiên làm việc
          try {
                if (Session["user"] == null)
                {
                    return Json(new { code = 401, }, JsonRequestBehavior.AllowGet);
                }
                return Json(new { code = 200, }, JsonRequestBehavior.AllowGet);
            } catch (Exception e) { 
                return Json(new { code = 500, msg = rm.GetString("false") + " " + e.Message }, JsonRequestBehavior.AllowGet); } 
        }
    }
}