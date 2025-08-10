using System.Globalization;
using ExcelEtlApp.Data;
using ExcelEtlApp.Models;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;

namespace ExcelEtlApp.Services
{
    public class ExcelEtlService
    {
        private readonly AppDbContext _db;
        private readonly ILogger<ExcelEtlService> _logger;

        public ExcelEtlService(AppDbContext db, ILogger<ExcelEtlService> logger)
        {
            _db = db;
            _logger = logger;
            // EPPlus license must be set in Program.cs before this service is used.
        }

        public async Task<EtlResult> RunEtlAsync(string filePath)
        {
            var result = new EtlResult();
            if (!File.Exists(filePath))
            {
                result.Errors.Add($"File not found: {filePath}");
                return result;
            }

            using var package = new ExcelPackage(new FileInfo(filePath));
            var workbook = package.Workbook;

            // 1. Parse lokasi sheet (sheet name: Lokasi)
            var lokasiSheet = workbook.Worksheets.FirstOrDefault(ws => string.Equals(ws.Name, "16", StringComparison.OrdinalIgnoreCase));
            if (lokasiSheet != null)
            {
                await ImportLokasi(lokasiSheet);
            }
            else
            {
                _logger.LogWarning("Sheet 'Lokasi' not found; lokasi lookup may fail.");
            }

            // 2. Parse TRX sheet (sheet name: TRX)
            var trxSheet = workbook.Worksheets.FirstOrDefault(ws => string.Equals(ws.Name, "16", StringComparison.OrdinalIgnoreCase));
            if (trxSheet != null)
            {
                var trxRows = ReadSheetForValues(trxSheet);
                foreach (var r in trxRows)
                {
                    // r.Location may be a lokasi name or id; try map to id
                    int idx = r.Location.Contains(".") ? r.Location.IndexOf(".") + 1 : 0;
                    string loc = r.Location.Substring(idx + 1, r.Location.Length - (idx + 1)).Trim();
                    var lokasiId = await FindLokasiIdAsync(loc);
                    if (lokasiId == null)
                    {
                        _logger.LogWarning($"Unknown lokasi '{r.Location}' in TRX; skipping row.");
                        continue;
                    }

                    var trx = new TRX
                    {
                        Bulan = new DateTime(r.Month.Year, r.Month.Month, 1),
                        LokasiId = lokasiId,
                        Trx = r.Value
                    };
                    _db.TRX.Add(trx);
                }
            }
            else
            {
                result.Warnings.Add("TRX sheet not found.");
            }

            // 3. Parse NPP sheet
            var nppSheet = workbook.Worksheets.FirstOrDefault(ws => string.Equals(ws.Name, "18", StringComparison.OrdinalIgnoreCase));
            if (nppSheet != null)
            {
                var nppRows = ReadSheetForValues(nppSheet);
                foreach (var r in nppRows)
                {
                    int idx = r.Location.Contains(".") ? r.Location.IndexOf(".") + 1 : 0;
                    string loc = r.Location.Substring(idx + 1, r.Location.Length - (idx + 1)).Trim();
                    var lokasiId = await FindLokasiIdAsync(loc);
                    if (lokasiId == null)
                    {
                        _logger.LogWarning($"Unknown lokasi '{r.Location}' in NPP; skipping row.");
                        continue;
                    }

                    var npp = new NPP
                    {
                        Bulan = new DateTime(r.Month.Year, r.Month.Month, 1),
                        LokasiId = lokasiId,
                        Npp = r.Value
                    };
                    _db.NPP.Add(npp);
                }
            }
            else
            {
                result.Warnings.Add("NPP sheet not found.");
            }

            await _db.SaveChangesAsync();
            result.Success = true;
            return result;
        }

        private async Task ImportLokasi(ExcelWorksheet ws)
        {
            // parse categories in column A like 'a. Jawa' and locations in column B
            string? currentPrefix = null;
            int counter = 0;
            int startRow = 4; // as in your sample
            for (int r = startRow; r <= ws.Dimension.End.Row; r++)
            {
                var a = ws.Cells[r, 1].Text?.Trim();
                var b = ws.Cells[r, 2].Text?.Trim();

                if (!string.IsNullOrEmpty(a) && a.Contains('.'))
                {
                    // new prefix like 'a.' -> 'a'
                    currentPrefix = a.Split('.')[0].Trim();
                    counter = 0;
                    continue;
                }

                if (!string.IsNullOrEmpty(b) && !string.IsNullOrEmpty(currentPrefix))
                {
                    counter++;
                    var id = $"{currentPrefix}{counter}";

                    var exists = await _db.Lokasi.AnyAsync(x => x.Id == id);
                    if (!exists)
                    {
                        int idx = b.Contains(".") ? b.IndexOf(".") + 1 : 0;
                        b = b.Substring(idx + 1, b.Length - (idx + 1)).Trim();
                        _db.Lokasi.Add(new Lokasi { Id = id, Nama = b });
                    }
                }
            }

            await _db.SaveChangesAsync();
        }

        private async Task<string?> FindLokasiIdAsync(string lokasiNameOrId)
        {
            if (string.IsNullOrWhiteSpace(lokasiNameOrId)) return null;

            // direct match to id
            var byId = await _db.Lokasi.FindAsync(lokasiNameOrId);
            if (byId != null) return byId.Id;

            // match by name (case-insensitive)
            var byName = await _db.Lokasi.FirstOrDefaultAsync(l => l.Nama.ToLower() == lokasiNameOrId.ToLower());
            if (byName != null) return byName.Id;

            return null;
        }

        private record SheetRow(string Location, DateTime Month, long Value);

        private List<SheetRow> ReadSheetForValues(ExcelWorksheet ws)
        {
            var list = new List<SheetRow>();
            // header row determination: assume row 3 has month headers
            int headerRow = 2;
            int firstDataRow = 4;
            int firstDataCol = 3; // months start at column C per your sample
            int lastCol = ws.Dimension.End.Column;
            int lastRow = ws.Dimension.End.Row;

            // read month headers from headerRow
            var months = new Dictionary<int, DateTime>();
            for (int c = firstDataCol; c <= lastCol; c++)
            {
                var header = ws.Cells[headerRow, c].Text?.Trim();
                if (string.IsNullOrEmpty(header)) continue;
                var parsed = ParseMonthFromHeader(header);
                months[c] = parsed;
            }

            for (int r = firstDataRow; r <= lastRow; r++)
            {
                var lokasiCell = ws.Cells[r, 2].Text?.Trim();
                if (string.IsNullOrEmpty(lokasiCell)) continue;
                if (lokasiCell.Equals("JUMLAH", StringComparison.OrdinalIgnoreCase)) break;

                for (int c = firstDataCol; c <= lastCol; c++)
                {
                    if (!months.ContainsKey(c)) continue;
                    var cell = ws.Cells[r, c].Text?.Trim();
                    if (string.IsNullOrEmpty(cell)) continue;
                    var val = ParseLongSafe(cell);
                    list.Add(new SheetRow(lokasiCell, months[c], val));
                }
            }

            return list;
        }

        private DateTime ParseMonthFromHeader(string header)
        {
            var tokens = header.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var t in tokens)
            {
                var s = t.Trim();
                if (s.Contains('/'))
                {
                    var parts = s.Split('/');
                    if (parts.Length == 2 && int.TryParse(parts[0], out int m) && int.TryParse(parts[1], out int y))
                    {
                        int year = y < 100 ? 2000 + y : y;
                        return new DateTime(year, m, 1);
                    }
                }
                if (DateTime.TryParseExact(s, new[] { "MMM yyyy", "MMMM yyyy", "MM/yyyy", "M/yyyy", "MM-yyyy", "M-yyyy", "MM-yy", "M-yy", "MMM-yy" }, CultureInfo.InvariantCulture, DateTimeStyles.None, out var dt))
                {
                    return new DateTime(dt.Year, dt.Month, 1);
                }
            }

            var now = DateTime.Now;
            return new DateTime(now.Year, now.Month, 1);
        }

        private long ParseLongSafe(string s)
        {
            var cleaned = new string(s.Where(c => char.IsDigit(c) || c == '-').ToArray());
            if (long.TryParse(cleaned, out var v)) return v;
            return 0;
        }
    }

    public class EtlResult
    {
        public bool Success { get; set; } = false;
        public List<string> Errors { get; } = new();
        public List<string> Warnings { get; } = new();
        public bool HasErrors => Errors.Any();
    }
}
