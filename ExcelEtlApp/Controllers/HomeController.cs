using ExcelEtlApp.Data;
using ExcelEtlApp.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;

namespace ExcelEtlApp.Controllers
{
    public class HomeController : Controller
    {
        private readonly AppDbContext _db;
        private readonly ILogger<HomeController> _logger;

        public HomeController(AppDbContext db, ILogger<HomeController> logger)
        {
            _db = db;
            _logger = logger;
        }
     
        public async Task<IActionResult> Index(string? lokasiId)
        {
            // list lokasi for dropdown
            var lokasiList = await _db.Lokasi.OrderBy(l => l.Id).ToListAsync();
            ViewBag.LokasiList = lokasiList;
            ViewBag.SelectedLokasi = lokasiId;

            List<MonthlyDto> data = new();

            if (!string.IsNullOrEmpty(lokasiId))
            {
                var tiketsize = await _db.TicketSizeView
                    .Where(n => n.lokasi == lokasiId).OrderBy(x=>x.bulan).ToListAsync();

                foreach (var m in tiketsize)
                {                    
                    data.Add(new MonthlyDto { Bulan = m.bulan, Trx = m.trx, Npp = m.npp, TicketSize = m.ticketsize });
                }
            }

            

            //var list = await q.OrderBy(t => t.lokasi).ToListAsync();
            return View(data);
        }

                
        public class MonthlyDto
        {
            public DateTime Bulan { get; set; }
            public long Trx { get; set; }
            public long Npp { get; set; }
            public double TicketSize { get; set; }
        }
    }
}
