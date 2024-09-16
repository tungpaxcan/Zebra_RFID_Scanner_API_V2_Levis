using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Resources;
using System.Web.Mvc;
using System.Web.UI.WebControls;
using Zebra_RFID_Scanner.Models;
using Newtonsoft.Json;
using OfficeOpenXml;
using System.Xml;
using System.Configuration;
using System.Net.Http;
using System.Threading.Tasks;
using CsvHelper;
using System.Globalization;
using System.Data;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.EMMA;
using DocumentFormat.OpenXml.InkML;
using DocumentFormat.OpenXml.VariantTypes;
using static Zebra_RFID_Scanner.Controllers.EPCTOUPC;
using System.Net;
using static Zebra_RFID_Scanner.Controllers.Ctn;

namespace Zebra_RFID_Scanner.Controllers
{
    public class FunctionOrderController : BaseController
    {

        private Entities db = new Entities();
        ResourceManager rm = new ResourceManager("Zebra_RFID_Scanner.Resources.Resource", typeof(Resources.Resource).Assembly);

        // GET: FunctionOrder
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult Rescan()
        {
            return View();
        }
        public ActionResult RescanAll()
        {
            return View();
        }
        public ActionResult RescanAll2()
        {
            return View();
        }
        public ActionResult RescanLevis2()
        {
            return View();
        }
        [HttpGet]
        public JsonResult list(string seach)
        {
            try
            {
                string[] removeFiles = Directory.GetFiles(Server.MapPath("/PreASN/"));
                for (int i = 0; i < removeFiles.Length; i++)
                {
                    System.IO.File.Delete(removeFiles[i]);
                }
                var itemsToDelete = db.Files.Where(item => item.Status == null).ToList();

                db.Files.RemoveRange(itemsToDelete);
                db.SaveChanges();
                // Copy the file with new name
                string folderPath = @"C:\RFID\PreASN"; // Đường dẫn đến folder
                string[] files = Directory.GetFiles(folderPath); // Lấy danh sách các file trong folder
                Array.Sort(files, (a, b) =>
          DateTime.Compare(new FileInfo(a).CreationTime, new FileInfo(b).CreationTime)
          );
                List<PreASN.PreAsn> pkls = new List<PreASN.PreAsn>();
                List<string> Po = new List<string>();
                List<string> Rsl = new List<string>();
                foreach (string file in files)
                {
                    var name = file.Substring(file.LastIndexOf('\\') + 1);
                    var extensionFile = file.Substring(file.LastIndexOf('.') + 1);
                    if (extensionFile != "xlsx")
                    {
                        Rsl.Add(name + ": " + rm.GetString("Thefilemusthavetheextensionxlsx").ToString());
                    }
                    else if (extensionFile == "xlsx")
                    {
                        using (var package = new ExcelPackage(new FileInfo(file)))
                        {
                            var worksheet = package.Workbook.Worksheets[0];
                            var noOfRow = worksheet.Dimension.End.Row;
                            //begin format
                            string header1 = worksheet.Cells[1, 1].Value?.ToString().Trim();
                            string header2 = worksheet.Cells[1, 2].Value?.ToString().Trim();
                            string header3 = worksheet.Cells[1, 3].Value?.ToString().Trim();
                            string header4 = worksheet.Cells[1, 4].Value?.ToString().Trim();
                            string header5 = worksheet.Cells[1, 5].Value?.ToString().Trim();
                            string header6 = worksheet.Cells[1, 6].Value?.ToString().Trim();
                            string header7 = worksheet.Cells[1, 7].Value?.ToString().Trim();
                            string header8 = worksheet.Cells[1, 8].Value?.ToString().Trim();
                            string header9 = worksheet.Cells[1, 9].Value?.ToString().Trim();
                            //identify header is RFID or UPC
                            if(header7 == "EPC-RFID")
                            {
                                if (header1 != "Factory") { Rsl.Add(name + ": " + rm.GetString("formisnotinthecorrectformat").ToString()); }
                                else if (header2 != "Fty Code") { Rsl.Add(name + ": " + rm.GetString("formisnotinthecorrectformat").ToString()); }
                                else if (header3 != "Port of Loading") { Rsl.Add(name + ": " + rm.GetString("formisnotinthecorrectformat").ToString()); }
                                else if (header4 != "PO") { Rsl.Add(name + ": " + rm.GetString("formisnotinthecorrectformat").ToString()); }
                                else if (header5 != "Ctn Label") { Rsl.Add(name + ": " + rm.GetString("formisnotinthecorrectformat").ToString()); }
                                else if (header6 != "UPC") { Rsl.Add(name + ": " + rm.GetString("formisnotinthecorrectformat").ToString()); }
                                else
                                {
                                    string po = worksheet.Cells[2, 4].Value.ToString();
                                    if (Po.Contains(po))
                                    {
                                        Rsl.Add("File " + name + " " + rm.GetString("coincide").ToString() + " PO " + po);
                                    }
                                    else
                                    {
                                        Po.Add(po);
                                        string sourceFilePath = @"C:\RFID\PreASN\" + name + "";
                                        string destinationFilePath = Server.MapPath("/PreASN/" + name + "");
                                        System.IO.File.Copy(sourceFilePath, destinationFilePath);
                                        PreASN.PreAsn pkl = new PreASN.PreAsn()
                                        {
                                            Path = "/PreASN/" + name + "",
                                            Name = name
                                        };
                                        pkls.Add(pkl);
                                        for (int i = 2; i <= noOfRow; i++)
                                        {
                                            var a = worksheet.Cells[i, 3].Value;
                                            if (a == null)
                                            {
                                                continue;
                                            }
                                            var filePo = worksheet.Cells[i, 4].Value.ToString();
                                            var fileCtn = worksheet.Cells[i, 5].Value.ToString();
                                            var fileUpc = worksheet.Cells[i, 6].Value.ToString();
                                            var checkFile = db.Files.SingleOrDefault(x =>x.Po == filePo && x.Ctn == fileCtn);
                                            if (checkFile == null)
                                            {
                                                Zebra_RFID_Scanner.Models.File file1 = new Zebra_RFID_Scanner.Models.File();
                                                file1.Po = filePo;
                                                file1.Ctn = fileCtn;
                                                file1.Upc = fileUpc;
                                                file1.CreateDate = DateTime.Now;
                                                db.Files.Add(file1);
                                                db.SaveChanges();
                                            }
                                        }
                                    }
                                }
                            }
                            else
                            {
                                header6 = header6 != "Carton to" ?  worksheet.Cells[1, 7].Value.ToString().Trim() : header6;
                                if (header1 != "Consignee") { Rsl.Add(name + ": " + rm.GetString("formisnotinthecorrectformat").ToString()); }
                                else if (header2 != "Shipper") { Rsl.Add(name + ": " + rm.GetString("formisnotinthecorrectformat").ToString()); }
                                else if (header3 != "SOnumber") { Rsl.Add(name + ": " + rm.GetString("formisnotinthecorrectformat").ToString()); }
                                else if (header4 != "PO") { Rsl.Add(name + ": " + rm.GetString("formisnotinthecorrectformat").ToString()); }
                                else if (header5 != "SKU") { Rsl.Add(name + ": " + rm.GetString("formisnotinthecorrectformat").ToString()); }
                                else if (header6 != "Carton to") { Rsl.Add(name + ": " + rm.GetString("formisnotinthecorrectformat").ToString()); }
                                else if (header8 != "UPC") { Rsl.Add(name + ": " + rm.GetString("formisnotinthecorrectformat").ToString()); }
                                else if (header9 != "UPC QTY") { Rsl.Add(name + ": " + rm.GetString("formisnotinthecorrectformat").ToString()); }
                                else
                                {
                                    string po = worksheet.Cells[2, 4].Value.ToString();
                                    string so = worksheet.Cells[2, 3].Value.ToString();
                                    if (Po.Contains(po))
                                    {
                                        Rsl.Add("File " + name + " " + rm.GetString("coincide").ToString() + " PO " + po);

                                    }
                                    else if (Po.Contains(so))
                                    {
                                        Rsl.Add("File " + name + " " + rm.GetString("coincide").ToString() + " SO " + so);
                                    }
                                    else
                                    {
                                        Po.Add(po);
                                        Po.Add(so);
                                        string sourceFilePath = @"C:\RFID\PreASN\" + name + "";
                                        string destinationFilePath = Server.MapPath("/PreASN/" + name + "");
                                        System.IO.File.Copy(sourceFilePath, destinationFilePath);
                                        PreASN.PreAsn pkl = new PreASN.PreAsn()
                                        {
                                            Path = "/PreASN/" + name + "",
                                            Name = name
                                        };
                                        pkls.Add(pkl);
                                        for (int i = 2; i <= noOfRow; i++)
                                        {
                                            var a = worksheet.Cells[i, 3].Value;
                                            if (a == null)
                                            {
                                                continue;
                                            }
                                            var fileSo = worksheet.Cells[i, 3].Value.ToString();
                                            var filePo = worksheet.Cells[i, 4].Value.ToString();
                                            var fileSku = worksheet.Cells[i, 5].Value.ToString();
                                            var fileCtn = worksheet.Cells[i, 6].Value.ToString();
                                            var fileUpc = worksheet.Cells[i, 8].Value.ToString();
                                            var fileQty = worksheet.Cells[i, 9].Value.ToString();
                                            var checkFile = db.Files.SingleOrDefault(x => x.So == fileSo && x.Po == filePo && x.Sku == fileSku && x.Ctn == fileCtn);
                                            if (checkFile == null)
                                            {
                                                Zebra_RFID_Scanner.Models.File file1 = new Zebra_RFID_Scanner.Models.File();
                                                file1.So = fileSo;
                                                file1.Po = filePo;
                                                file1.Sku = fileSku;
                                                file1.Ctn = fileCtn;
                                                file1.Upc = fileUpc;
                                                file1.Qty = fileQty;
                                                file1.CreateDate = DateTime.Now;
                                                db.Files.Add(file1);
                                                db.SaveChanges();
                                            }
                                        }
                                    }
                                }
                            }


                        }
                    }
                }
                if (seach != "")
                {
                    var pklsSeach = pkls.ToList().Where(x => x.Name.Contains(seach) || x.Name.ToLower().Contains(seach));
                    return Json(new { code = 200, pkl = pklsSeach }, JsonRequestBehavior.AllowGet);
                }
                else
                {
                    var pklsSeach = pkls.ToList().Where(x => x.Name.Contains(seach) || x.Name.ToLower().Contains(seach));
                    return Json(new { code = 200, pkl = pklsSeach, rsl = Rsl }, JsonRequestBehavior.AllowGet);
                }

            }
            catch (Exception e)
            {
                return Json(new { code = 500, msg = "Sai !!!" + e.Message }, JsonRequestBehavior.AllowGet);
            }
        }
        [HttpPost]
        public JsonResult DeleteEPC()
        {
            try
            {
                var session = (User)Session["user"];
                var FX = session.FXconnect.Id;
                Dele.DeleteDetailEPCs(FX);
                return Json(new { code = 200, msg = rm.GetString("success").ToString() }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception e)
            {
                return Json(new { code = 500, msg = rm.GetString("false").ToString() + e.Message }, JsonRequestBehavior.AllowGet);
            }
        }

        [HttpGet]
        public JsonResult CtnWrong(string ctn)
        {
            try
            {
                var po = db.Files.SingleOrDefault(x => x.Ctn == ctn);
                if (po != null)
                {
                    return Json(new { code = 200, po = po }, JsonRequestBehavior.AllowGet);
                }
                else
                {
                    return Json(new { code = 300, }, JsonRequestBehavior.AllowGet);
                }

            }
            catch (Exception e)
            {
                return Json(new { code = 500, msg = rm.GetString("false") + " " + e.Message }, JsonRequestBehavior.AllowGet);
            }
        }


    }
}