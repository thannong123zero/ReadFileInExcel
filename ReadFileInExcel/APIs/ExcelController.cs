using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;

namespace ReadFileInExcel.APIs
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelController : ControllerBase
    {
        private readonly IWebHostEnvironment _webHostEnvironment;
        public ExcelController(IWebHostEnvironment webHostEnvironment)
        {
            _webHostEnvironment = webHostEnvironment;
        }
        [HttpGet]
        public async Task<IActionResult> Get()
        {
            string imageUrl = "https://images2.thanhnien.vn/528068263637045248/2024/1/25/c3c8177f2e6142e8c4885dbff89eb92a-65a11aeea03da880-1706156293184503262817.jpg";
            // Validate the URL to prevent misuse (optional, but recommended)
            if (!Uri.IsWellFormedUriString(imageUrl, UriKind.Absolute))
            {
                return BadRequest("Invalid image URL.");
            }
            ExtractPdfFromExcel();
            using (var httpClient = new HttpClient())
            {
                // Fetch the image from the external server
                var response = await httpClient.GetAsync(imageUrl);
                if (!response.IsSuccessStatusCode)
                {
                    return NotFound("Image not found.");
                }

                // Read the image content
                var content = await response.Content.ReadAsByteArrayAsync();
                var contentType = response.Content.Headers.ContentType.ToString();


                // Return the image as a FileResult
                return File(content, contentType);
            }
        }
        [HttpPost]
        public IActionResult Post([FromForm] IFormFile file)
        {
            if (file == null || file.Length == 0)
                return BadRequest("No file uploaded.");

            var path = Path.Combine(_webHostEnvironment.WebRootPath, "uploads");
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            string filePath = Path.Combine(path, file.FileName);
            using (var stream = new FileStream(filePath, FileMode.Create))
            {
                file.CopyTo(stream);
            }

            // Process the Excel file to extract the embedded PDF
            //var pdfPath = ExtractPdfFromExcel(path);

            return Ok();
            
        }
        private void ExtractPdfFromExcel()
        {
            var path = Path.Combine(_webHostEnvironment.WebRootPath, "uploads");
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            var filePath = Path.Combine(_webHostEnvironment.WebRootPath, "uploads", "book1.xlsx");

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(filePath))
            {
                var worksheet = package.Workbook.Worksheets[0];
                var rowCount = worksheet.Dimension.Rows;

                // Map images from the worksheet drawings
                var drawings = worksheet.Drawings;

                for (int row = 2; row <= rowCount; row++) // Start from row 2 (skip headers)
                {
                    var record = new ExcelRecord();

                    // Safely extract cell values
                    record.STT = worksheet.Cells[row, 1].GetValue<int>();
                    record.EquipmentType = worksheet.Cells[row, 2].GetValue<string>();
                    record.Unit = worksheet.Cells[row, 3].GetValue<string>();
                    record.Name = worksheet.Cells[row, 4].GetValue<string>();
                    record.Quantity = worksheet.Cells[row, 5].GetValue<string>();

                    record.Note = worksheet.Cells[row, 8].GetValue<string>();
                    
                    // Extract image for the current row
                    foreach (var drawing in drawings)
                    {
                        //if (drawing is ExcelPicture picture)
                        //{
                        //    // Match the picture with the row based on position
                        //    if (picture.From.Row + 1 == row) // EPPlus row index starts from 0
                        //    {
                        //        string imageName = $"image_{row - 1}.png";
                        //        string imagePath = Path.Combine(path, imageName);

                        //        using (var imageStream = new FileStream(imagePath, FileMode.Create))
                        //        {
                        //            imageStream.Write(picture.Image.ImageBytes, 0, picture.Image.ImageBytes.Length);
                        //        }
                        //    }
                        //}
                       
                    }
                }
            }
        }
        

        [HttpPost("uploadImage")]
        public async Task<IActionResult> UploadImage([FromForm] IFormFile image)
        {
            string path = Path.Combine(_webHostEnvironment.WebRootPath, "assets", "images");
            if (image == null || image.Length == 0)
            {
                return BadRequest("Please select an image file.");
            }
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            string imageName = $"image_{DateTime.Now.ToString("yyyyMMddHHmmss")}.png";
            string imagePath = Path.Combine(path, imageName);
            using (var stream = new FileStream(imagePath, FileMode.Create))
            {
                await image.CopyToAsync(stream);
            }

            return Ok(image);
        }
    }
}
public class ExcelRecord
{
    public int STT { get; set; }
    public string EquipmentType { get; set; }
    public string Unit { get; set; }
    public string Name { get; set; }
    public string Quantity { get; set; }
    public string ImageFile { get; set; }
    public string File { get; set; }
    public string Note { get; set; } // Add this property to fix the error
}
