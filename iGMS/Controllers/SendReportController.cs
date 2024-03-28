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

namespace Zebra_RFID_Scanner.Controllers
{
    public class SendReportController : BaseController
    {
        Entities db = new Entities();
        ReportsController reportsController = new ReportsController();
        // GET: SendReport
        [HttpPost]
        public JsonResult SendGeneral(string id, string name, string hod)
        {
            try
            {
                DateTime dateTime = DateTime.Now;   
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
                        C5 = "SoNum",
                        C6 = "doNo",
                        C7 = "setCd",
                        C8 = "subDoNo",
                        C9 = "mngFctryCd",
                        C10 = "facBranchCd",
                        C11 = "shipperCode",
                        C12 = "packKey",
                        C13 = "expectedQty",
                        C14 = "qty",
                        C15 = "missingEpc",
                        C16 = "unknownEpc",
                        C17 = "uploadDateTime"
                    };
                    datas.Add(data);
                    var reports = db.Reports.SingleOrDefault(x => x.Id == id);
                    if(reports == null) { return Json(new { code = 500, msg = "There are no files to report !!!" }, JsonRequestBehavior.AllowGet); }
                    var list = db.Generals.Where(x => x.IdReports == id).ToList();
                    foreach (var item in list)
                    {
                        ReportsEPC datum = new ReportsEPC()
                        {
                            C1 = reports.TimeStart == null ? "" : reports.TimeStart.Value.Year + "-" + reports.TimeStart.Value.Month.ToString().PadLeft(2, '0') + "-" + reports.TimeStart.Value.Day.ToString().PadLeft(2, '0'),
                            C2 = item.cntry,
                            C3 = item.port,
                            C4 = item.deviceNum,
                            C5 = item.So,
                            C6 = item.doNo,
                            C7 = item.setCd,
                            C8 = item.subDoNo,
                            C9 = item.mngFctryCd,
                            C10 = item.facBranchCd,
                            C11 = reports.Shiper,
                            C12 = item.packKey,
                            C13 = item.Qty,
                            C14 = item.Qty,
                            C15 = "0",
                            C16 = "Matched",
                            C17 = dateTime.Year.ToString().PadLeft(2,'0') + "-" + dateTime.Month.ToString().PadLeft(2, '0') + "-" + dateTime.Day.ToString().PadLeft(2, '0')+" "+dateTime.Hour.ToString().PadLeft(2,'0')+":"+
                                    dateTime.Minute.ToString().PadLeft(2, '0')+":"+dateTime.Second.ToString().PadLeft(2, '0')+":"+dateTime.Millisecond.ToString().PadLeft(3,'0')
                        };
                        datas.Add(datum);
                    }
                    var listDis = db.Discrepancies.Where(x => x.IdReports == id).ToList();
                    foreach (var item in listDis)
                    {
                        ReportsEPC datum = new ReportsEPC()
                        {
                            C1 = reports.TimeStart == null ? "" : reports.TimeStart.Value.Year + "-" + reports.TimeStart.Value.Month.ToString().PadLeft(2, '0') + "-" + reports.TimeStart.Value.Day.ToString().PadLeft(2, '0'),
                            C2 = item.cntry,
                            C3 = item.port,
                            C4 = item.deviceNum,
                            C5 = item.So,
                            C6 = item.doNo,
                            C7 = item.setCd,
                            C8 = item.subDoNo,
                            C9 = item.mngFctryCd,
                            C10 = item.facBranchCd,
                            C11 = reports.Shiper,
                            C12 = item.packKey,
                            C13 = item.Qty,
                            C14 = item.QtyScan,
                            C15 = Math.Abs(int.Parse(item.QtyScan)-int.Parse(item.Qty)).ToString(),
                            C16 = "Mismatched",
                            C17 = dateTime.Year.ToString().PadLeft(2, '0') + "-" + dateTime.Month.ToString().PadLeft(2, '0') + "-" + dateTime.Day.ToString().PadLeft(2, '0') + " " + dateTime.Hour.ToString().PadLeft(2, '0') + ":" +
                                    dateTime.Minute.ToString().PadLeft(2, '0') + ":" + dateTime.Second.ToString().PadLeft(2, '0') + ":" + dateTime.Millisecond.ToString().PadLeft(3, '0')
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
                            for (int j = 0; j < 17; j++)
                            {
                                var propertyName = "C" + (j + 1);
                                var property = typeof(ReportsEPC).GetProperty(propertyName);
                                var propertyValue = property.GetValue(datas[i]);
                                
                                worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].Value = propertyValue;
                                worksheet.Cells[GetExcelColumnName(j + 1) + (i + 1)].AutoFitColumns();
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