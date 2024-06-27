﻿using System;
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

namespace Zebra_RFID_Scanner.Controllers
{
    public class FunctionOrderController : BaseController
    {

        private Entities db = new Entities();
        ResourceManager rm = new ResourceManager("Zebra_RFID_Scanner.Resources.Resource", typeof(Resources.Resource).Assembly);

        // GET: FunctionOrder
        public ActionResult Index()
        {
            try
            {
                var xmlDelete = db.XMLValidates.Where(x => x.Id > 0).ToList();
                db.XMLValidates.RemoveRange(xmlDelete);
                db.SaveChanges();
                string folderPath = @"C:\RFID\ScanFile"; // Đường dẫn đến folder
                string[] files = Directory.GetFiles(folderPath); // Lấy danh sách các file trong folder
                Array.Sort(files, (a, b) =>
          DateTime.Compare(new FileInfo(a).CreationTime, new FileInfo(b).CreationTime)
          );
                foreach (string file in files)
                {
                    var extensionFile = file.Substring(file.LastIndexOf('.') + 1);
                    if (extensionFile == "xml")
                    {
                        XmlDocument xmlDoc = new XmlDocument();
                        xmlDoc.Load(file);
                        XmlNodeList details = xmlDoc.SelectNodes("//SOLINEDETAIL");
                        foreach (XmlNode detail in details)
                        {
                            XMLValidate validate = new XMLValidate();
                            foreach (XmlNode header in detail.ParentNode.ChildNodes)
                            {
                                if (header.LocalName == "PO")
                                {
                                    validate.Po = header.InnerText;
                                }
                                if (header.LocalName == "SO")
                                {
                                    validate.So = header.InnerText;
                                }
                                if (header.LocalName == "SKU")
                                {
                                    validate.Sku = header.InnerText;
                                    continue;
                                }
                            }
                            foreach (XmlNode detailchil in detail.ChildNodes)
                            {
                                if (detailchil.LocalName == "CARTONID")
                                {
                                    validate.Ctn = detailchil.InnerText;
                                    continue;
                                }
                            }
                            validate.Status = true;
                            db.XMLValidates.Add(validate);
                            db.SaveChanges();
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);

            }
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
                    if (extensionFile != "xlsx" && extensionFile != "csv")
                    {
                        Rsl.Add(name + ": " + rm.GetString("Thefilemusthavetheextensionxlsx").ToString());
                    }
                    else if (extensionFile == "xlsx")
                    {
                        using (var package = new ExcelPackage(new FileInfo(file)))
                        {
                            var worksheet = package.Workbook.Worksheets["Sheet1"];
                            var noOfRow = worksheet.Dimension.End.Row;
                            //begin format

                            string consignee = worksheet.Cells[1, 1].Value.ToString().Trim();
                            string shipper = worksheet.Cells[1, 2].Value.ToString().Trim();
                            string sonumber = worksheet.Cells[1, 3].Value.ToString().Trim();
                            string ponumber = worksheet.Cells[1, 4].Value.ToString().Trim();
                            string sku = worksheet.Cells[1, 5].Value.ToString().Trim();
                            string cartonto = worksheet.Cells[1, 6].Value.ToString().Trim() != "Carton to" ? worksheet.Cells[1, 7].Value.ToString().Trim() : worksheet.Cells[1, 6].Value.ToString().Trim();
                            string upc = worksheet.Cells[1, 8].Value.ToString().Trim();
                            string upcqty = worksheet.Cells[1, 9].Value.ToString().Trim();

                            if (consignee != "Consignee") { Rsl.Add(name + ": " + rm.GetString("formisnotinthecorrectformat").ToString()); }
                            else if (shipper != "Shipper") { Rsl.Add(name + ": " + rm.GetString("formisnotinthecorrectformat").ToString()); }
                            else if (sonumber != "SOnumber") { Rsl.Add(name + ": " + rm.GetString("formisnotinthecorrectformat").ToString()); }
                            else if (ponumber != "PO") { Rsl.Add(name + ": " + rm.GetString("formisnotinthecorrectformat").ToString()); }
                            else if (sku != "SKU") { Rsl.Add(name + ": " + rm.GetString("formisnotinthecorrectformat").ToString()); }
                            else if (cartonto != "Carton to" && cartonto != "EPC") { Rsl.Add(name + ": " + rm.GetString("formisnotinthecorrectformat").ToString()); }
                            else if (upc != "UPC") { Rsl.Add(name + ": " + rm.GetString("formisnotinthecorrectformat").ToString()); }
                            else if (upcqty != "UPC QTY" && upcqty != "Scan Datetime") { Rsl.Add(name + ": " + rm.GetString("formisnotinthecorrectformat").ToString()); }
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
                    //else if (extensionFile == "csv")
                    //{
                    //    using (var reader = new StreamReader(file))
                    //    using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
                    //    {
                    //        var records = csv.GetRecords<dynamic>().ToList();
                    //        var recordDict = (IDictionary<string, object>)records[0];

                    //        string doNo = recordDict.Keys.ToArray()[0].ToString().Trim();
                    //        string subDoNo = recordDict.Keys.ToArray()[1].ToString().Trim();
                    //        string mngFctryCd = recordDict.Keys.ToArray()[2].ToString().Trim();
                    //        string facBranchCd = recordDict.Keys.ToArray()[3].ToString().Trim();
                    //        string shipperCode = recordDict.Keys.ToArray()[4].ToString().Trim();
                    //        string setCd = recordDict.Keys.ToArray()[5].ToString().Trim();
                    //        string cntNo = recordDict.Keys.ToArray()[6].ToString().Trim();
                    //        string yr = recordDict.Keys.ToArray()[7].ToString().Trim();
                    //        string ssnCd = recordDict.Keys.ToArray()[8].ToString().Trim();
                    //        string dptPortCd = recordDict.Keys.ToArray()[9].ToString().Trim();
                    //        string cntry = recordDict.Keys.ToArray()[10].ToString().Trim();
                    //        string exf = recordDict.Keys.ToArray()[11].ToString().Trim();
                    //        string packKey = recordDict.Keys.ToArray()[12].ToString().Trim();
                    //        string epc = recordDict.Keys.ToArray()[13].ToString().Trim();

                    //        if (doNo != "doNo") { Rsl.Add(name + ": " + rm.GetString("formisnotinthecorrectformat").ToString() + " doNo"); }
                    //        else if (subDoNo != "subDoNo") { Rsl.Add(name + ": " + rm.GetString("formisnotinthecorrectformat").ToString() + " subDoNo"); }
                    //        else if (mngFctryCd != "mngFctryCd") { Rsl.Add(name + ": " + rm.GetString("formisnotinthecorrectformat").ToString() + " mngFctryCd"); }
                    //        else if (facBranchCd != "facBranchCd") { Rsl.Add(name + ": " + rm.GetString("formisnotinthecorrectformat").ToString() + " facBranchCd"); }
                    //        else if (shipperCode != "shipperCode") { Rsl.Add(name + ": " + rm.GetString("formisnotinthecorrectformat").ToString() + " shipperCode"); }
                    //        else if (setCd != "setCd") { Rsl.Add(name + ": " + rm.GetString("formisnotinthecorrectformat").ToString() + " setCd"); }
                    //        else if (cntNo != "cntNo") { Rsl.Add(name + ": " + rm.GetString("formisnotinthecorrectformat").ToString() + " cntNo"); }
                    //        else if (yr != "yr") { Rsl.Add(name + ": " + rm.GetString("formisnotinthecorrectformat").ToString() + " yr"); }
                    //        else if (ssnCd != "ssnCd") { Rsl.Add(name + ": " + rm.GetString("formisnotinthecorrectformat").ToString() + " ssnCd"); }
                    //        else if (dptPortCd != "dptPortCd") { Rsl.Add(name + ": " + rm.GetString("formisnotinthecorrectformat").ToString() + " dptPortCd"); }
                    //        else if (cntry != "cntry") { Rsl.Add(name + ": " + rm.GetString("formisnotinthecorrectformat").ToString() + " cntry"); }
                    //        else if (exf != "exf") { Rsl.Add(name + ": " + rm.GetString("formisnotinthecorrectformat").ToString() + " exf"); }
                    //        else if (packKey != "packKey") { Rsl.Add(name + ": " + rm.GetString("formisnotinthecorrectformat").ToString() + " packKey"); }
                    //        else if (epc != "epc") { Rsl.Add(name + ": " + rm.GetString("formisnotinthecorrectformat").ToString() + " epc"); }
                    //        else
                    //        {
                    //            string sourceFilePath = @"C:\RFID\PreASN\" + name + "";
                    //            string destinationFilePath = Server.MapPath("/PreASN/" + name + "");
                    //            System.IO.File.Copy(sourceFilePath, destinationFilePath);
                    //            PreASN.PreAsn pkl = new PreASN.PreAsn()
                    //            {
                    //                Path = "/PreASN/" + name + "",
                    //                Name = name
                    //            };
                    //            pkls.Add(pkl);
                    //        }
                    //    }
                    //}
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
        ApiController ApiController = new ApiController();
        [HttpPost]
        public async Task<JsonResult> GetFileApi()
        {
            try
            {
                
                string hostApi = ConfigurationManager.ConnectionStrings["HostApi"].ConnectionString;
                string port_code = Request.Form["portCode"];
                string date = Request.Form["date"];
                string device = Request.Form["device"];
                string[] PO = new string[0];
                if (string.IsNullOrEmpty(date) || string.IsNullOrEmpty(port_code))
                {
                    return Json(new { code = 500, msg = rm.GetString("false").ToString() + " Not enough input to get data" }, JsonRequestBehavior.AllowGet);
                }
                var checkReport = db.Reports.FirstOrDefault(x => x.Id == port_code + "_" +date);
                if (checkReport != null)
                {
                    if(device == "pc")
                    {
                        return Json(new { code = 500, msg = "The file already exists in the system" }, JsonRequestBehavior.AllowGet);
                    }
                    else
                    {
                        PO = checkReport.Po.Split(',');
                        return Json(new { code = 200, PO, TypeStatus = true}, JsonRequestBehavior.AllowGet);
                    }
                    
                }
                Certificate2 cert = await ApiController.GetCertificate();
                if (cert != null)
                {
                    if (cert.cert != null)
                    {
                        var handler = new HttpClientHandler();
                        handler.ClientCertificates.Add(cert.cert);
                        var client = new HttpClient(handler);

                        var request = new HttpRequestMessage()
                        {
                            RequestUri = new Uri($"{hostApi}/tps/presign?action=download&port_code={port_code}&date={date}"),
                            Method = HttpMethod.Get,
                        };
                        var response = await client.SendAsync(request);
                        if (response.IsSuccessStatusCode)
                        {
                            var responsecontent = await response.Content.ReadAsAsync<string>();
                            ReponseApi jsonResponse = JsonConvert.DeserializeObject<ReponseApi>(responsecontent);

                            if (jsonResponse.StatusCode == 200)
                            {
                                string folderName = @"C:\RFID\PreASN\";
                                string fileUrl = port_code + "_" + date + ".csv";
                                await DownloadFileAsync(jsonResponse.Body, folderName, fileUrl);
                                return Json(new { code = 200, msg = "Retrieve File successfully", fileUrl,PO, TypeStatus = false }, JsonRequestBehavior.AllowGet);
                            }
                            else if (jsonResponse.StatusCode == 400)
                            {
                                return Json(new { code = 500, msg = rm.GetString("false").ToString() + " " + jsonResponse.Body }, JsonRequestBehavior.AllowGet);
                            }
                            else
                            {
                                return Json(new { code = 500, msg = rm.GetString("false").ToString() + " " + jsonResponse.Body }, JsonRequestBehavior.AllowGet);
                            }

                        }
                        else
                        {
                            return Json(new { code = 500, msg = rm.GetString("false").ToString() + " Cannot get data, wrong request" }, JsonRequestBehavior.AllowGet);
                        }
                    }
                    else
                    {
                        return Json(new { code = 500, msg = rm.GetString("false").ToString() + " " + cert.msg }, JsonRequestBehavior.AllowGet);
                    }
                }
                else
                {
                    return Json(new { code = 500, msg = rm.GetString("false").ToString() + " Certificate not installed" }, JsonRequestBehavior.AllowGet);
                }
            }
            catch (Exception ex)
            {
                return Json(new { code = 500, msg = rm.GetString("false").ToString() + " " + ex.Message+ex.InnerException }, JsonRequestBehavior.AllowGet);
            }

        }
        ReportsController controller = new ReportsController();
        HomeController Hcontroller = new HomeController();

        private async Task DownloadFileAsync(string fileUrl, string destinationFolder, string namePath)
        {
            try
            {
               
                List<List<string>> rows = new List<List<string>>();
                Certificate2 cert = await ApiController.GetCertificate();
                if (cert != null)
                {
                    if (cert.cert != null)
                    {
                        var handler = new HttpClientHandler();
                        handler.ClientCertificates.Add(cert.cert);
                        var client = new HttpClient(handler);

                        var request = new HttpRequestMessage()
                        {
                            RequestUri = new Uri(fileUrl),
                            Method = HttpMethod.Get,
                        };
                        var response = await client.SendAsync(request);
                        response.EnsureSuccessStatusCode();
                        if (response.IsSuccessStatusCode)
                        {
                            List<Discrepancy> general = new List<Discrepancy>();
                            List<EPCDiscrepancy> ePCDiscrepancies = new List<EPCDiscrepancy>();
                            string Po = "";
                            using (var stream = await response.Content.ReadAsStreamAsync())
                            using (var reader = new StreamReader(stream))
                            {

                                List<string> rowValuesHead = new List<string>();
                                rowValuesHead.Add("doNo");
                                rowValuesHead.Add("subDoNo");
                                rowValuesHead.Add("mngFctryCd");
                                rowValuesHead.Add("facBranchCd");
                                rowValuesHead.Add("shipperCode");
                                rowValuesHead.Add("setCd");
                                rowValuesHead.Add("cntNo");
                                rowValuesHead.Add("yr");
                                rowValuesHead.Add("ssnCd");
                                rowValuesHead.Add("dptPortCd");
                                rowValuesHead.Add("cntry");
                                rowValuesHead.Add("exf");
                                rowValuesHead.Add("packKey");
                                rowValuesHead.Add("epc");
                                rowValuesHead.Add("UPC");
                                rows.Add(rowValuesHead);
                                while (!reader.EndOfStream)
                                {
                                    var line = await reader.ReadLineAsync();
                                    var values = line.Split(',');
                                    List<string> rowValues = new List<string>(values);
                                    rowValues.Add("s");
                                    // Thêm giá trị tính toán từ controller.epctoupc() vào cột thứ 14 (index 13)
                                    string computedValue = controller.epctoupc(rowValues[13]);
                                    string CartonTo = rowValues[6]?.ToString().Trim();
                                    string EPC = rowValues[13]?.ToString().Trim();
                                    string doNo = rowValues[0]?.ToString().Trim();
                                    var checkD = general.FirstOrDefault(x=>x.CartonTo==CartonTo&&x.Po == doNo);
                                    var checkE = ePCDiscrepancies.FirstOrDefault(x=>x.Id== EPC);
                                    if (!Po.Contains(doNo)) { Po += "," + doNo; };
                                    if (checkD==null) {
                                        Discrepancy d = new Discrepancy()
                                        {
                                            So = "",
                                            Po = doNo,
                                            CartonTo = CartonTo,
                                            UPC = CartonTo,
                                            Qty = "1",
                                            Sku = "",
                                            Status = false,
                                            QtyScan = "0",
                                            cntry = rowValues[10]?.ToString().Trim(),
                                            port = rowValues[9]?.ToString().Trim(),
                                            deviceNum = "",
                                            doNo = doNo,
                                            setCd = rowValues[5]?.ToString().Trim(),
                                            subDoNo = rowValues[1]?.ToString().Trim(),
                                            mngFctryCd = rowValues[2]?.ToString().Trim(),
                                            facBranchCd = rowValues[3]?.ToString().Trim(),
                                            packKey = rowValues[12]?.ToString().Trim(),
                                        };
                                        general.Add(d);
                                    }
                                    else
                                    {
                                        checkD.Qty = (int.Parse(checkD.QtyScan) + 1).ToString();
                                    }
                                    if(checkE == null)
                                    {
                                        EPCDiscrepancy ePC = new EPCDiscrepancy()
                                        {
                                            Carton = CartonTo,
                                            Id = EPC,
                                            So = "",
                                            Po = doNo,
                                            Sku = "",
                                            UPC = CartonTo,
                                            cntry = rowValues[10]?.ToString().Trim(),
                                            port = rowValues[9]?.ToString().Trim(),
                                            deviceNum = "",
                                            doNo = doNo,
                                            setCd = rowValues[5]?.ToString().Trim(),
                                            subDoNo = rowValues[1]?.ToString().Trim(),
                                            mngFctryCd = rowValues[2]?.ToString().Trim(),
                                            facBranchCd = rowValues[3]?.ToString().Trim(),
                                            packKey = rowValues[12]?.ToString().Trim(),
                                        };
                                        ePCDiscrepancies.Add(ePC);
                                    }
                                   
                                    
                                    rowValues.Add(computedValue);
                                    // Thêm hàng mới vào danh sách các hàng
                                    rows.Add(rowValues);
                                }
                            }
                            string Ctn = JsonConvert.SerializeObject(general);
                            string EPCDiscrepancy = JsonConvert.SerializeObject(ePCDiscrepancies);
                            Hcontroller.Save(namePath.Replace(".csv",""), "", Ctn, "", Po, "", "", "", "", "", "", EPCDiscrepancy);
                            // Ghi dữ liệu đã được mở rộng vào file CSV
                            using (var fileStream = new FileStream(destinationFolder + Path.GetFileName(namePath), FileMode.Create))
                            using (var writer = new StreamWriter(fileStream))
                            {
                                foreach (var row in rows)
                                {
                                    await writer.WriteLineAsync(string.Join(",", row));
                                }
                             
                            }
                            string destinationFilePath = Server.MapPath("/PreASN/" + namePath + "");
                            System.IO.File.Copy(destinationFolder + Path.GetFileName(namePath), destinationFilePath);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
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