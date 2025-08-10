using ExcelEtlApp.Models;
using Microsoft.EntityFrameworkCore;

namespace ExcelEtlApp.Data
{
    public class AppDbContext : DbContext
    {
        public AppDbContext(DbContextOptions<AppDbContext> options) : base(options) { }

        public DbSet<Lokasi> Lokasi { get; set; } = null!;
        public DbSet<TRX> TRX { get; set; } = null!;
        public DbSet<NPP> NPP { get; set; } = null!;

        public DbSet<TicketSizeView> TicketSizeView { get; set; }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            // map to lowercase table names to avoid Postgres case-sensitivity issues
            modelBuilder.Entity<Lokasi>().ToTable("lokasi");
            modelBuilder.Entity<TRX>().ToTable("trx");
            modelBuilder.Entity<NPP>().ToTable("npp");
            modelBuilder.Entity<TicketSizeView>().HasNoKey().ToView("ticket_size");
        

            modelBuilder.Entity<Lokasi>().HasKey(l => l.Id);
            modelBuilder.Entity<TRX>().HasIndex(r => new { r.Bulan, r.LokasiId });
            modelBuilder.Entity<NPP>().HasIndex(r => new { r.Bulan, r.LokasiId });

            // ensure Bulan stored as date (no time)
            modelBuilder.Entity<TRX>().Property(p => p.Bulan).HasColumnType("date");
            modelBuilder.Entity<NPP>().Property(p => p.Bulan).HasColumnType("date");
        }
    }
}
