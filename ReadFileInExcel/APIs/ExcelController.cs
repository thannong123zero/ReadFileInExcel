using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml.Drawing;
using OfficeOpenXml;

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

                FileContentResult fileContent = new FileContentResult(content,contentType);
                // Return the image as a FileResult
                return Ok(fileContent);
            }
        }
        [HttpPost]
        public async Task<IActionResult> Post([FromForm] IFormFile excelFile)
        {
            string path = Path.Combine(_webHostEnvironment.WebRootPath,"assets","images");
            if (excelFile == null || excelFile.Length == 0)
            {
                return BadRequest("Please select an Excel file.");
            }
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            var records = new List<ExcelRecord>();

            using (var stream = new MemoryStream())
            {
                await excelFile.CopyToAsync(stream);
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var package = new ExcelPackage(stream))
                {
                    var worksheet = package.Workbook.Worksheets[0]; // Read the first worksheet
                    var rowCount = worksheet.Dimension.Rows;

                    for (int row = 2; row <= rowCount; row++) // Assuming the first row has headers
                    {
                        var record = new ExcelRecord();
                        record.STT = int.Parse(worksheet.Cells[row, 1].Text);
                        record.CardId = worksheet.Cells[row, 2].Text;
                        record.Fullname = worksheet.Cells[row, 3].Text;


                        foreach (var drawing in worksheet.Drawings)
                        {
                            if (drawing is ExcelPicture picture)
                            {
                                // Check if the image is positioned near the relevant row
                                if (picture.From.Row + 1 == row) // Adjust column index if necessary
                                {
                                    // Extract the image bytes
                                    var imageBytes = picture.Image.ImageBytes;

                                    // Create an in-memory IFormFile
                                    var image = new FormFile(
                                        new MemoryStream(imageBytes),
                                        0,
                                        imageBytes.Length,
                                        "image",
                                        picture.Name)
                                    {
                                        Headers = new HeaderDictionary(),
                                        ContentType = "image/jpeg" // Adjust content type based on the image format
                                    };

                                    // Use the image (e.g., assign to a model property)
                                    record.File = image;

                                    break; // Stop after finding the relevant image for the row
                                }
                            }
                        }

                        records.Add(record);
                    }
                }
            }
            return Ok(records);
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
    public string CardId { get; set; }
    public string Fullname { get; set; }
    public DateTime BirthDay { get; set; }
    public string PhoneNumber { get; set; }
    public IFormFile File { get; set; }
    public string Note { get; set; }
}