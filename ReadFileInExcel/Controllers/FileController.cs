using System.IO;
using System.Linq;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Syncfusion.XlsIO;
using iTextSharp.text.pdf;

public class FileController : Controller
{
    private readonly IWebHostEnvironment _environment;

    public FileController(IWebHostEnvironment environment)
    {
        _environment = environment;
    }

    [HttpPost]
    public IActionResult UploadFile(IFormFile file)
    {
        if (file == null || file.Length == 0)
            return BadRequest("No file uploaded.");

        var filePath = Path.Combine(_environment.WebRootPath, "uploads", file.FileName);
        using (var stream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
        {
            file.CopyTo(stream);
        }

        // Process the Excel file to extract the embedded PDF
        var pdfPath = ExtractPdfFromExcel(filePath);

        return Ok(new { PdfPath = pdfPath });
    }

    private string ExtractPdfFromExcel(string filePath)
    {
        // Use Syncfusion.XlsIO to read the Excel file and extract the PDF
        using (ExcelEngine excelEngine = new ExcelEngine())
        {
            IApplication application = excelEngine.Excel;
            // Open the workbook from a file stream
            using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                IWorkbook workbook = application.Workbooks.Open(fileStream);
                IWorksheet worksheet = workbook.Worksheets[0];

                // Assuming the PDF is embedded in a specific cell as a string
                var cell = worksheet["A1"];
                var pdfBase64 = cell.DisplayText;

                if (!string.IsNullOrEmpty(pdfBase64))
                {
                    // Convert the base64 string to byte array
                    var pdfBytes = Convert.FromBase64String(pdfBase64);
                    var pdfPath = Path.Combine(_environment.WebRootPath, "uploads", "extracted.pdf");
                    System.IO.File.WriteAllBytes(pdfPath, pdfBytes);
                    return pdfPath;
                }
            }
        }

        return string.Empty; // Return an empty string instead of null
    }
}
