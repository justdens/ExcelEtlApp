using ExcelEtlApp.Services;
using Microsoft.AspNetCore.Mvc;

namespace ExcelEtlApp.Controllers
{
    public class UploadController : Controller
    {
        private readonly ExcelEtlService _etl;
        private readonly IWebHostEnvironment _env;

        public UploadController(ExcelEtlService etl, IWebHostEnvironment env)
        {
            _etl = etl;
            _env = env;
        }

        [HttpGet]
        public IActionResult Index() => View();

        [HttpPost]
        public async Task<IActionResult> Index(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                ViewBag.Message = "Please select a file.";
                return View();
            }

            var path = Path.Combine(_env.ContentRootPath, "wwwroot", "data.xlsx");
            using (var stream = new FileStream(path, FileMode.Create))
            {
                await file.CopyToAsync(stream);
            }

            var res = await _etl.RunEtlAsync(path);
            ViewBag.Success = res.Success;
            ViewBag.Errors = res.Errors;
            ViewBag.Warnings = res.Warnings;
            return View();
        }
    }
}
