using OfficeOpenXml.Style;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Zebra_RFID_Scanner.Models;
using Newtonsoft.Json;
using System.Text;
using System.Net.Http;
using System.Threading.Tasks;
using System.Net.Http.Headers;
using System.Configuration;
using System.Resources;

namespace Zebra_RFID_Scanner.Controllers
{
    public class SendReportController : BaseController
    {
        Entities db = new Entities();
        ReportsController reportsController = new ReportsController();
        ResourceManager rm = new ResourceManager("Zebra_RFID_Scanner.Resources.Resource", typeof(Resources.Resource).Assembly);
        // GET: SendReport
        [HttpPost]
        public async Task<JsonResult> SendGeneral(string id, string name, string hod,string type)
        {
            try
            {
                DateTime dateTime = DateTime.Now;
                var Port = "";
                string checkFileString = JsonConvert.SerializeObject(reportsController.rescanAll(id).Data);
                ReponseCheckFile checkFile = JsonConvert.DeserializeObject<ReponseCheckFile>(checkFileString);
                if(checkFile.status) 
                {
                    hod = hod.Substring(0, hod.IndexOf(" "));
                    List<ReportsEPC> datas = new List<ReportsEPC>();
                    ReportsEPC data = new ReportsEPC()
                    {
                        C1 = "scanDate",
                        C2 = "cntry",
                        C3 = "port",
                        C4 = "deviceNum",
                        //C5 = "SoNum",
                        C5 = "doNo",
                        C6 = "setCd",
                        C7 = "subDoNo",
                        C8 = "mngFctryCd",
                        C9 = "facBranchCd",
                        C10 = "shipperCode",
                        C11 = "packKey",
                        C12 = "expectedQty",
                        C13 = "qty",
                        C14 = "missingEpc",
                        C15 = "epcMoveToOtherCartons",
                        C16 = "unknownEpc",
                        C17 = "epcMoveFromOtherCartons",
                        C18 = "uploadDateTime"
                    };
                    datas.Add(data);
                    var reports = db.Reports.SingleOrDefault(x => x.Id == id);
                    if(reports == null) { return Json(new { code = 500, msg = "There are no files to report !!!" }, JsonRequestBehavior.AllowGet); }
                    var list = db.Generals.Where(x => x.IdReports == id).ToList();
                    foreach (var item in list)
                    {
                        ReportsEPC datum = new ReportsEPC()
                        {
                            C1 = reports.TimeStart == null ? "" : reports.TimeStart.Value.Month.ToString().PadLeft(2, '0') + "/" + reports.TimeStart.Value.Day.ToString().PadLeft(2, '0') + "/" + reports.TimeStart.Value.Year ,
                            C2 = item.cntry,
                            C3 = item.port,
                            C4 = item.deviceNum,
                            //C5 = item.So,
                            C5 = item.doNo,
                            C6 = item.setCd,
                            C7 = item.subDoNo,
                            C8 = item.mngFctryCd,
                            C9 = item.facBranchCd,
                            C10 = reports.Shiper,
                            C11 = item.packKey,
                            C12 = item.Qty,
                            C13 = item.Qty,
                            C14 = "0",
                            C15 = "0",
                            C16 = "0",
                            C17 = "0",
                            C18 = dateTime.Year.ToString().PadLeft(2,'0') + "-" + dateTime.Month.ToString().PadLeft(2, '0') + "-" + dateTime.Day.ToString().PadLeft(2, '0')+" "+dateTime.Hour.ToString().PadLeft(2,'0')+":"+
                                    dateTime.Minute.ToString().PadLeft(2, '0')+":"+dateTime.Second.ToString().PadLeft(2, '0')+":"+dateTime.Millisecond.ToString().PadLeft(3,'0')
                        };
                        Port = item.port;
                        datas.Add(datum);
                    }
                    var listDis = db.Discrepancies.Where(x => x.IdReports == id).ToList();
                    foreach (var item in listDis)
                    {
                        ReportsEPC datum = new ReportsEPC()
                        {
                            C1 = reports.TimeStart == null ? "" :  reports.TimeStart.Value.Month.ToString().PadLeft(2, '0') + "/" + reports.TimeStart.Value.Day.ToString().PadLeft(2, '0') + "/" + reports.TimeStart.Value.Year ,
                            C2 = item.cntry,
                            C3 = item.port,
                            C4 = item.deviceNum,
                            C5 = item.doNo,
                            C6 = item.setCd,
                            C7 = item.subDoNo,
                            C8 = item.mngFctryCd,
                            C19 = item.facBranchCd,
                            C10 = reports.Shiper,
                            C11 = item.packKey,
                            C12 = item.Qty,
                            C13 = item.QtyScan,
                            C14 = (int.Parse(item.QtyScan) - int.Parse(item.Qty)<0? Math.Abs(int.Parse(item.QtyScan) - int.Parse(item.Qty)).ToString() : "0"),
                            C15 = (int.Parse(item.QtyScan) - int.Parse(item.Qty) < 0 ? Math.Abs(int.Parse(item.QtyScan) - int.Parse(item.Qty)).ToString() : "0"),
                            C16 = (int.Parse(item.QtyScan) - int.Parse(item.Qty) > 0 ? Math.Abs(int.Parse(item.QtyScan) - int.Parse(item.Qty)).ToString() : "0"),
                            C17 = (int.Parse(item.QtyScan) - int.Parse(item.Qty) > 0 ? Math.Abs(int.Parse(item.QtyScan) - int.Parse(item.Qty)).ToString() : "0"),
                            C18 = dateTime.Year.ToString().PadLeft(2, '0') + "-" + dateTime.Month.ToString().PadLeft(2, '0') + "-" + dateTime.Day.ToString().PadLeft(2, '0') + " " + dateTime.Hour.ToString().PadLeft(2, '0') + ":" +
                                    dateTime.Minute.ToString().PadLeft(2, '0') + ":" + dateTime.Second.ToString().PadLeft(2, '0') + ":" + dateTime.Millisecond.ToString().PadLeft(3, '0')
                        };
                        Port = item.port;
                        datas.Add(datum);
                    }
                    string folderName = @"C:\RFID\Reports\";
                    string filePath = Port +"_"+ dateTime.Year.ToString().PadLeft(2, '0') + "-" + dateTime.Month.ToString().PadLeft(2, '0') + "-" + dateTime.Day.ToString().PadLeft(2, '0') + ".csv";
                    string folderPath = folderName + filePath;
                    string folderPathLocal = folderName + id+".csv";
                    using (StreamWriter writer = new StreamWriter(folderPathLocal, false, Encoding.UTF8))
                    {
                      

                        // Ghi dữ liệu vào file CSV
                        for (int i = 0; i < datas.Count; i++)
                        {
                            for (int j = 0; j < 18; j++)
                            {
                                if (j > 0)
                                {
                                    writer.Write(",");
                                }
                                var propertyName = "C" + (j + 1);
                                var property = typeof(ReportsEPC).GetProperty(propertyName);
                                var propertyValue = property.GetValue(datas[i]);

                                if (propertyValue != null)
                                {
                                    writer.Write(propertyValue.ToString().Replace(",", ";")); // Thay thế dấu phẩy để tránh lỗi CSV
                                }
                            }
                            writer.WriteLine();
                        }
                    }
                    if(type == "send")
                    {
                        await UploadFileApi(Port, dateTime.Year.ToString().PadLeft(2, '0') + "-" + dateTime.Month.ToString().PadLeft(2, '0') + "-" + dateTime.Day.ToString().PadLeft(2, '0'), folderPathLocal);
                    }
                    return Json(new { code = 200, }, JsonRequestBehavior.AllowGet);
                }
                else
                {
                    return Json(new { code = 500, msg = "Cannot send this file !!!" }, JsonRequestBehavior.AllowGet);
                }
               
            }
            catch (Exception e)
            {
                return Json(new { code = 500, msg = " !!!" + e.Message }, JsonRequestBehavior.AllowGet);
            }
        }
        ApiController ApiController = new ApiController();
        [HttpPost]
        public async Task<JsonResult> UploadFileApi(string port_code, string date,string folderPathLocal)
        {
            try
            {
                string hostApi = ConfigurationManager.ConnectionStrings["HostApi"].ConnectionString;
                if (string.IsNullOrEmpty(date) || string.IsNullOrEmpty(port_code))
                {
                    return Json(new { code = 500, msg = rm.GetString("false").ToString() + " Not enough input to get data" }, JsonRequestBehavior.AllowGet);
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
                            RequestUri = new Uri($"{hostApi}/tps/presign?action=upload&port_code={port_code}&date={date}"),
                            Method = HttpMethod.Get,
                        };
                        var response = await client.SendAsync(request);
                        if (response.IsSuccessStatusCode)
                        {
                            var responsecontent = await response.Content.ReadAsAsync<string>();
                            ReponseApi jsonResponse = JsonConvert.DeserializeObject<ReponseApi>(responsecontent);

                            if (jsonResponse.StatusCode == 200)
                            {
                                await UploadFileToApi(folderPathLocal, jsonResponse.Body);
                                return Json(new { code = 200, msg = "Upload File successfully" }, JsonRequestBehavior.AllowGet);
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
                return Json(new { code = 500, msg = rm.GetString("false").ToString() + " " + ex.Message }, JsonRequestBehavior.AllowGet);
            }

        }
        public async Task UploadFileToApi(string folderPathLocal, string apiUrl)
        {
            Certificate2 cert = await ApiController.GetCertificate();
            if (cert != null)
            {
                if (cert.cert != null)
                {
                    var handler = new HttpClientHandler();
                    handler.ClientCertificates.Add(cert.cert);
                    var client = new HttpClient(handler);

                    using (var content = new MultipartFormDataContent())
                    {
                        // Sử dụng Stream để đọc file CSV
                        using (FileStream fileStream = new FileStream(folderPathLocal, FileMode.Open, FileAccess.Read))
                        {
                            var fileContent = new StreamContent(fileStream);
                            fileContent.Headers.ContentType = MediaTypeHeaderValue.Parse("multipart/form-data");

                            // Thêm file CSV vào nội dung của yêu cầu HTTP
                            content.Add(fileContent, "file", Path.GetFileName(folderPathLocal));
                            // Gửi yêu cầu HTTP POST
                            var request = new HttpRequestMessage()
                            {
                                RequestUri = new Uri(apiUrl),
                                Method = HttpMethod.Put,
                                Content = fileContent
                            };
                            var response = await client.SendAsync(request);
                            response.EnsureSuccessStatusCode();
                            // Kiểm tra kết quả
                            if (response.IsSuccessStatusCode)
                            {
                                Console.WriteLine("File uploaded successfully.");
                                var responsecontent = await response.Content.ReadAsAsync<string>();
                                ReponseApi jsonResponse = JsonConvert.DeserializeObject<ReponseApi>(responsecontent);
                            }
                            else
                            {
                                Console.WriteLine($"Failed to upload file. Status code: {response.StatusCode}");
                            }
                        }
                    }
                }
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
    }
}