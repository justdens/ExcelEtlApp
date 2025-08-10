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

            // 1. Import Lokasi dari sheet TRX (kolom A & B)
            var trxSheet = workbook.Worksheets.FirstOrDefault(ws => ws.Name.Trim() == "16");
            if (trxSheet != null)
            {
                await ImportLokasi(trxSheet);
            }
            else
            {
                result.Warnings.Add("Sheet '16' untuk TRX tidak ditemukan.");
            }

            // 2. Proses TRX
            if (trxSheet != null)
            {
                var trxRows = ReadSheetForValues(trxSheet);
                var trxList = new List<TRX>();

                foreach (var r in trxRows)
                {
                    var lokasiId = await FindLokasiIdAsync(ExtractLocationName(r.Location));
                    if (lokasiId == null)
                    {
                        result.Warnings.Add($"Unknown lokasi '{r.Location}' di TRX; skipping.");
                        continue;
                    }

                    trxList.Add(new TRX
                    {
                        Bulan = new DateTime(r.Month.Year, r.Month.Month, 1),
                        LokasiId = lokasiId,
                        Trx = r.Value
                    });
                }

                var (validTrx, trxWarnings) = ValidateTrx(trxList);
                result.Warnings.AddRange(trxWarnings);
                _db.TRX.AddRange(validTrx);
            }

            // 3. Proses NPP
            var nppSheet = workbook.Worksheets.FirstOrDefault(ws => ws.Name.Trim() == "18");
            if (nppSheet != null)
            {
                var nppRows = ReadSheetForValues(nppSheet);
                var nppList = new List<NPP>();

                foreach (var r in nppRows)
                {
                    var lokasiId = await FindLokasiIdAsync(ExtractLocationName(r.Location));
                    if (lokasiId == null)
                    {
                        result.Warnings.Add($"Unknown lokasi '{r.Location}' di NPP; skipping.");
                        continue;
                    }

                    nppList.Add(new NPP
                    {
                        Bulan = new DateTime(r.Month.Year, r.Month.Month, 1),
                        LokasiId = lokasiId,
                        Npp = r.Value
                    });
                }

                var (validNpp, nppWarnings) = ValidateNpp(nppList);
                result.Warnings.AddRange(nppWarnings);
                _db.NPP.AddRange(validNpp);
            }
            else
            {
                result.Warnings.Add("Sheet '18' untuk NPP tidak ditemukan.");
            }

            await _db.SaveChangesAsync();
            result.Success = true;
            return result;
        }

        private async Task ImportLokasi(ExcelWorksheet ws)
        {
            string? currentPrefix = null;
            int counter = 0;
            int startRow = 4;

            for (int r = startRow; r <= ws.Dimension.End.Row; r++)
            {
                var a = ws.Cells[r, 1].Text?.Trim();
                var b = ws.Cells[r, 2].Text?.Trim();

                if (!string.IsNullOrEmpty(a) && a.Contains('.'))
                {
                    currentPrefix = a.Split('.')[0].Trim();
                    counter = 0;
                    continue;
                }

                if (!string.IsNullOrEmpty(b) && !string.IsNullOrEmpty(currentPrefix))
                {
                    counter++;
                    var id = $"{currentPrefix}{counter}";

                    if (!await _db.Lokasi.AnyAsync(x => x.Id == id))
                    {
                        _db.Lokasi.Add(new Lokasi { Id = id, Nama = ExtractLocationName(b) });
                    }
                }
            }

            await _db.SaveChangesAsync();
        }

        private async Task<string?> FindLokasiIdAsync(string lokasiNameOrId)
        {
            if (string.IsNullOrWhiteSpace(lokasiNameOrId)) return null;

            var byId = await _db.Lokasi.FindAsync(lokasiNameOrId);
            if (byId != null) return byId.Id;

            var byName = await _db.Lokasi.FirstOrDefaultAsync(l => l.Nama.ToLower() == lokasiNameOrId.ToLower());
            return byName?.Id;
        }

        private string ExtractLocationName(string raw)
        {
            int idx = raw.Contains(".") ? raw.IndexOf(".") + 1 : 0;
            return raw.Substring(idx).Trim();
        }

        private record SheetRow(string Location, DateTime Month, long Value);

        private List<SheetRow> ReadSheetForValues(ExcelWorksheet ws)
        {
            var list = new List<SheetRow>();
            int headerRow = 2;
            int firstDataRow = 4;
            int firstDataCol = 3;
            int lastCol = ws.Dimension.End.Column;
            int lastRow = ws.Dimension.End.Row;

            var months = new Dictionary<int, DateTime>();
            for (int c = firstDataCol; c <= lastCol; c++)
            {
                var header = ws.Cells[headerRow, c].Text?.Trim();
                if (!string.IsNullOrEmpty(header))
                    months[c] = ParseMonthFromHeader(header);
            }

            for (int r = firstDataRow; r <= lastRow; r++)
            {
                var lokasiCell = ws.Cells[r, 2].Text?.Trim();
                if (string.IsNullOrEmpty(lokasiCell) || lokasiCell.Equals("JUMLAH", StringComparison.OrdinalIgnoreCase))
                    continue;

                for (int c = firstDataCol; c <= lastCol; c++)
                {
                    if (!months.ContainsKey(c)) continue;
                    var valText = ws.Cells[r, c].Text?.Trim();
                    if (string.IsNullOrEmpty(valText)) continue;
                    list.Add(new SheetRow(lokasiCell, months[c], ParseLongSafe(valText)));
                }
            }

            return list;
        }

        private DateTime ParseMonthFromHeader(string header)
        {
            var tokens = header.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var s in tokens)
            {
                if (s.Contains('/'))
                {
                    var parts = s.Split('/');
                    if (parts.Length == 2 && int.TryParse(parts[0], out int m) && int.TryParse(parts[1], out int y))
                    {
                        int year = y < 100 ? 2000 + y : y;
                        return new DateTime(year, m, 1);
                    }
                }
                if (DateTime.TryParseExact(s,
                    new[] { "MMM yyyy", "MMMM yyyy", "MM/yyyy", "M/yyyy", "MM-yyyy", "M-yyyy", "MM-yy", "M-yy", "MMM-yy" },
                    CultureInfo.InvariantCulture, DateTimeStyles.None, out var dt))
                {
                    return new DateTime(dt.Year, dt.Month, 1);
                }
            }
            return new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
        }

        private long ParseLongSafe(string s)
        {
            var cleaned = new string(s.Where(c => char.IsDigit(c) || c == '-').ToArray());
            return long.TryParse(cleaned, out var v) ? v : 0;
        }

        private (List<TRX>, List<string>) ValidateTrx(List<TRX> records)
        {
            var logs = new List<string>();
            var valid = new List<TRX>();

            // LEFT JOIN ke Lokasi untuk ambil nama
            var recordsWithName = from r in records
                                  join l in _db.Lokasi on r.LokasiId equals l.Id into lj
                                  from l in lj.DefaultIfEmpty()
                                  select new { Rec = r, LokasiName = l != null ? l.Nama : r.LokasiId };

            var groups = recordsWithName.GroupBy(x => x.LokasiName);
            foreach (var g in groups)
            {
                var ordered = g.OrderBy(x => x.Rec.Bulan).ToList();
                TRX? prev = null;
                foreach (var item in ordered)
                {
                    var rec = item.Rec;
                    if (prev != null && rec.Trx <= prev.Trx)
                    {
                        prev = rec;
                        logs.Add($"TRX validation failed for {item.LokasiName} at {rec.Bulan:yyyy-MM}: {rec.Trx} <= {prev.Trx}");
                        continue;
                    }
                    valid.Add(rec);
                    prev = rec;
                }
            }

            return (valid, logs);
        }

        private (List<NPP>, List<string>) ValidateNpp(List<NPP> records)
        {
            var logs = new List<string>();
            var valid = new List<NPP>();

            var recordsWithName = from r in records
                                  join l in _db.Lokasi on r.LokasiId equals l.Id into lj
                                  from l in lj.DefaultIfEmpty()
                                  select new { Rec = r, LokasiName = l != null ? l.Nama : r.LokasiId };

            var groups = recordsWithName.GroupBy(x => x.LokasiName);
            foreach (var g in groups)
            {
                var ordered = g.OrderBy(x => x.Rec.Bulan).ToList();
                NPP? prev = null;
                foreach (var item in ordered)
                {
                    var rec = item.Rec;
                    if (prev != null && rec.Npp < prev.Npp)
                    {
                        prev = rec;
                        logs.Add($"NPP validation failed for {item.LokasiName} at {rec.Bulan:yyyy-MM}: {rec.Npp} < {prev.Npp}");
                        continue;
                    }
                    valid.Add(rec);
                    prev = rec;
                }
            }

            return (valid, logs);
        }
    }

    public class EtlResult
    {
        public bool Success { get; set; }
        public List<string> Errors { get; } = new();
        public List<string> Warnings { get; } = new();
        public bool HasErrors => Errors.Any();
    }
}
