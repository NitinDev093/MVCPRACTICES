using ClosedXML.Excel;
using MVCPRACTICES.Utility;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Web;
using System.Web.Mvc;

namespace MVCPRACTICES.Controllers
{
    public class BulkUploadController : Controller
    {
        // GET: BulkUpload
        public ActionResult UploadFile()
        {
            return View();
        }

        public ActionResult DownloadTemplate()
        {
            string TemplateName = "EmployeeTemplate";
            List<string> columns = new List<string>();
            columns = BulkUploadColumns.Employee;
            ExcelPackage licenseContext = new ExcelPackage();
            using (XLWorkbook wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Sheet1");
                for (int i = 0; i < columns.Count; i++)
                {
                    var cell = ws.Cell(1, i + 1); // Row 1, Column starts from 1
                    cell.Value = columns[i];

                    cell.Style.Font.Bold = true;
                    cell.Style.Fill.BackgroundColor = XLColor.LightBlue;
                    cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                    cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                }

                ws.Columns().AdjustToContents();
                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string fileName = $"{TemplateName}_{timestamp}.xlsx";
                using (MemoryStream stream = new MemoryStream())
                {
                    wb.SaveAs(stream);
                    return File(
                        stream.ToArray(),
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        fileName
                    );
                }
            }
        }


    [HttpPost]
    public JsonResult UploadExcel(HttpPostedFileBase excelFile)
    {
            try
            {

                if (excelFile == null || excelFile.ContentLength == 0)
                {
                    return Json(new { success = false, message = "No file selected." });
                }

                using (var package = new ExcelPackage(excelFile.InputStream))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    int totalRows = worksheet.Dimension.Rows;
                    int totalCols = worksheet.Dimension.Columns;

                    var data = new List<object>();

                    for (int row = 2; row <= totalRows; row++) // Assuming first row is header
                    {
                        data.Add(new
                        {
                            Name = worksheet.Cells[row, 1].Text,
                            Email = worksheet.Cells[row, 2].Text,
                            Phone = worksheet.Cells[row, 3].Text
                        });
                    }

                    // Optional: Save data to database here
                    return Json(new { success = true, message = "Excel processed successfully", data = data });
                }
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }

    }
}