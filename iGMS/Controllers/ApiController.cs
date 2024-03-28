using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeOpenXml.Style;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Newtonsoft.Json;
using System.Configuration;
using System.Net.Http;
using System.Security.Cryptography.X509Certificates;
using System.Threading.Tasks;
using Zebra_RFID_Scanner.Models;

namespace Zebra_RFID_Scanner.Controllers
{
    public class ApiController : Controller
    {
        // GET: Api
        [HttpPost]
        public JsonResult PushFile(List<DataApi> data, string nameFile)
        {
            try
            {
                List<ReportsEPC> datas = new List<ReportsEPC>();
                ReportsEPC header = new ReportsEPC()
                {
                    C1 = "PO",
                    C2 = "Ctn barcode",
                    C3 = "EPC",
                    C4 = "UPC",

                };
                datas.Add(header);
                foreach (var item in data)
                {
                    ReportsEPC datum = new ReportsEPC()
                    {
                        C1 = item.PO,
                        C2 = item.Ctn,
                        C3 = item.EPC,
                        C4 = item.UPC,
                    };
                    datas.Add(datum);
                }
                string folderName = @"C:\RFID\PreASN\";
                string filePath = nameFile + ".xlsx";
                string folderPath = folderName + filePath;
                using (ExcelPackage excelPackage = new ExcelPackage())
                {
                    // Add a new worksheet to the workbook
                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");

                    // Add some content to the worksheet
                    for (int i = 0; i < datas.Count; i++)
                    {
                        for (int j = 0; j < 4; j++)
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
                return Json(new { code = 200, msg = "Thành công đẩy file " + nameFile + ".xlsx" }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception e)
            {
                return Json(new { code = 500, msg = "Lỗi Ngoại Lệ" + e.Message }, JsonRequestBehavior.AllowGet);
            }
        }

        public async Task<Certificate2> GetCertificate()
        {
            try
            {
                var nameCertificate = "";
                var nameCertificatePass = "";

                // Đường dẫn đến thư mục Certificate
                string directoryPath = System.Web.Hosting.HostingEnvironment.MapPath("/Certificate");
                // Lấy danh sách tập tin .pfx và .p12 trong thư mục
                string[] certificateFiles = Directory.GetFiles(directoryPath, "*.pfx");
                string[] p12Files = Directory.GetFiles(directoryPath, "*.p12");
                // Lấy danh sách tập tin .txt trong thư mục
                string[] txtFiles = Directory.GetFiles(directoryPath, "*.txt");
                // Kết hợp hai mảng thành một mảng duy nhất
                List<string> allCertificateFiles = new List<string>(certificateFiles);
                allCertificateFiles.AddRange(p12Files);
                string[] allFiles = allCertificateFiles.ToArray();
                // Kiểm tra xem có tập tin nào không
                if (allFiles.Length > 0)
                {
                    foreach (string filePath in allFiles)
                    {
                        // Lấy tên tập tin từ đường dẫn hoàn chỉnh
                        string fileName = Path.GetFileName(filePath);
                        nameCertificate = fileName;
                    }
                }
                else
                {
                    return new Certificate2
                    {
                        msg = "There are no certificate files in the directory Certificate",
                        cert = null,
                    };
                }
                // Kiểm tra xem có tập tin txt nào không
                if (txtFiles.Length > 0)
                {
                    if (txtFiles.Length == 1)
                    {
                        foreach (string filePath in txtFiles)
                        {
                            // Lấy tên tập tin từ đường dẫn hoàn chỉnh
                            nameCertificatePass = System.IO.File.ReadAllText(filePath);
                        }
                    }
                    else
                    {
                        return new Certificate2
                        {
                            msg = "Cannot read certificate password because there are too many files",
                            cert = null,
                        };

                    }
                }
                else
                {
                    return new Certificate2
                    {
                        msg = "There is no certificate password",
                        cert = null,
                    };
                }

                var certPath = System.Web.Hosting.HostingEnvironment.MapPath($"/Certificate/{nameCertificate}");
                var cert = new X509Certificate2(certPath, nameCertificatePass);
                return new Certificate2
                {
                    msg = "Success",
                    cert = cert,
                };
            }
            catch (Exception ex)
            {
                return new Certificate2
                {
                    msg = ex.Message,
                    cert = null,
                };
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