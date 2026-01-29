using Microsoft.EntityFrameworkCore;
using PlanMaster.Models;

namespace PlanMaster.Data;

public class PlanDbContext : DbContext
{
    private readonly string _dbPath;

    // Новый набор: планы
    public DbSet<Plan> Plans => Set<Plan>();

    // Таблицы плана
    public DbSet<PlanTable> Tables => Set<PlanTable>();
    public DbSet<PlanRow> Rows => Set<PlanRow>();

    // Итоговая таблица плана
    public DbSet<SummaryTable> SummaryTables => Set<SummaryTable>();
    public DbSet<SummaryRow> SummaryRows => Set<SummaryRow>();

    public PlanDbContext(string dbPath)
    {
        _dbPath = dbPath;
    }

    protected override void OnConfiguring(DbContextOptionsBuilder options)
        => options.UseSqlite($"Data Source={_dbPath}");

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        modelBuilder.Entity<PlanTable>()
            .HasIndex(t => new { t.PlanId, t.SheetName });

        modelBuilder.Entity<Plan>()
            .HasMany<PlanTable>()
            .WithOne()
            .HasForeignKey(t => t.PlanId)
            .OnDelete(DeleteBehavior.Cascade);

        modelBuilder.Entity<Plan>()
            .HasOne<SummaryTable>()
            .WithOne()
            .HasForeignKey<SummaryTable>(s => s.PlanId)
            .OnDelete(DeleteBehavior.Cascade);

        modelBuilder.Entity<PlanTable>()
            .HasMany(t => t.Rows)
            .WithOne(r => r.PlanTable!)
            .HasForeignKey(r => r.PlanTableId)
            .OnDelete(DeleteBehavior.Cascade);

        modelBuilder.Entity<SummaryTable>()
            .HasMany(t => t.Rows)
            .WithOne(r => r.SummaryTable!)
            .HasForeignKey(r => r.SummaryTableId)
            .OnDelete(DeleteBehavior.Cascade);

        modelBuilder.Entity<PlanRow>()
            .HasIndex(r => new { r.PlanTableId, r.RowOrder });

        modelBuilder.Entity<SummaryRow>()
            .HasIndex(r => new { r.SummaryTableId, r.RowOrder });
    }
}