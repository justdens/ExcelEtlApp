using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace ExcelEtlApp.Models
{
    [Table("lokasi")]
    public class Lokasi
    {
        [Key]
        [Column("id")]
        public string Id { get; set; } = string.Empty;

        [Column("lokasi")]
        public string Nama { get; set; } = string.Empty;
    }
}
