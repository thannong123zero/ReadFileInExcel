using System.Diagnostics;
using Microsoft.AspNetCore.Mvc;
using ReadFileInExcel.Models;

namespace ReadFileInExcel.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly IWebHostEnvironment _hostingEnvironment;
        public HomeController(ILogger<HomeController> logger,IWebHostEnvironment webHostEnvironment)
        {
            _logger = logger;
            _hostingEnvironment = webHostEnvironment;
        }

        public IActionResult Index()
        {
            string path = Path.Combine(_hostingEnvironment.WebRootPath, "assets", "images", "image_20241129131341.png");
            byte[] imageBytes = System.IO.File.ReadAllBytes(path);
            string encodedImage = Convert.ToBase64String(imageBytes);

            ViewBag.EncodedImage = encodedImage;
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
