using System.ComponentModel.DataAnnotations.Schema;

namespace ExcelEtlApp.Models
{
    [NotMapped]
    public class TicketSizeView
    {
        public string lokasiid { get; set; } = string.Empty;
        public string lokasi { get; set; } = string.Empty;
        public DateTime bulan { get; set; }
        public long trx { get; set; }
        public long npp { get; set; }
        public double ticketsize { get; set; }
    }
}
