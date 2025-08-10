using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace ExcelEtlApp.Models
{
    [Table("npp")]
    public class NPP
    {
        [Key]
        [Column("id")]
        public int Id { get; set; }

        [Column("bulan")]
        public DateTime Bulan { get; set; }

        [Column("lokasi_id")]
        public string LokasiId { get; set; } = string.Empty;

        [Column("npp")]
        public long Npp { get; set; }
    }
}
