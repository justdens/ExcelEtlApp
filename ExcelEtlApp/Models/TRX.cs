using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace ExcelEtlApp.Models
{
    [Table("trx")]
    public class TRX
    {
        [Key]
        [Column("id")]
        public int Id { get; set; }

        [Column("bulan")]
        public DateTime Bulan { get; set; }

        [Column("lokasi_id")]
        public string LokasiId { get; set; } = string.Empty;

        [Column("trx")]
        public long Trx { get; set; }
    }
}
