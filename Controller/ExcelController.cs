using System.IO;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;

namespace ExcelWriter.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class ExcelController : ControllerBase
    {
        [HttpPost]
        [Route("update")]
        public IActionResult UpdateExcel([FromForm] string filepath, [FromForm] string text, [FromForm] int row, [FromForm] string column)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            FileInfo fileInfo = new FileInfo(filepath);

            using (var package = new ExcelPackage(fileInfo))
            {
                if (package.Workbook.Worksheets.Count == 0)
                {
                    package.Workbook.Worksheets.Add("Sheet1");
                }

                var worksheet = package.Workbook.Worksheets[0];

                // Convert row and column to Excel cell address
                string cellAddress = $"{column.ToUpper()}{row}";

                // Check if the column in the current row is not empty
                while (!string.IsNullOrEmpty(worksheet.Cells[cellAddress].Text))
                {
                    row++;
                    // Convert row and column to Excel cell address
                     cellAddress = $"{column.ToUpper()}{row}";
                }
                cellAddress = $"{column.ToUpper()}{row}";
                // Update cell with new text
                worksheet.Cells[cellAddress].Value = text;

                package.Save();
            }

            return Ok("Text updated in Excel successfully!");
        }
    }
}
