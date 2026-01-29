using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.EntityFrameworkCore;
using PlanMaster.Data;
using PlanMaster.Models;

namespace PlanMaster.Services;

public class PlanRepository
{
    private readonly string _dbPath;

    private const string DraftPlanName = "Черновик";

    public PlanRepository(string dbPath)
    {
        _dbPath = dbPath;
        using var db = new PlanDbContext(_dbPath);
        db.Database.EnsureCreated();
    }

    private PlanDbContext CreateContext() => new PlanDbContext(_dbPath);

    // -----------------------------
    //  A) Совместимость со старым UI
    //     SaveAll/LoadAll работают как "Черновик"
    // -----------------------------

    public async Task SaveAllAsync(List<PlanTable> tables, SummaryTable summary)
    {
        var draftId = await EnsureDraftPlanAsync();
        await SavePlanAsync(draftId, DraftPlanName, tables, summary);
    }

    public async Task<(List<PlanTable> Tables, SummaryTable? Summary)> LoadAllAsync()
    {
        var draftId = await EnsureDraftPlanAsync();
        return await LoadPlanAsync(draftId);
    }

    private async Task<int> EnsureDraftPlanAsync()
    {
        await using var db = CreateContext();

        var draft = await db.Plans.FirstOrDefaultAsync(p => p.Name == DraftPlanName);
        if (draft != null)
            return draft.Id;

        draft = new Plan
        {
            Name = DraftPlanName,
            CreatedAtUtc = DateTime.UtcNow,
            UpdatedAtUtc = DateTime.UtcNow
        };

        db.Plans.Add(draft);
        await db.SaveChangesAsync();

        return draft.Id;
    }

    // -----------------------------
    //  B) Планы
    // -----------------------------

    public async Task<List<Plan>> ListPlansAsync()
    {
        await using var db = CreateContext();
        return await db.Plans
            .OrderByDescending(p => p.UpdatedAtUtc)
            .ToListAsync();
    }

    public async Task<int> SaveAsNewPlanAsync(string? name, List<PlanTable> tables, SummaryTable summary)
    {
        await using var db = CreateContext();

        var plan = new Plan
        {
            Name = string.IsNullOrWhiteSpace(name)
                ? $"План {DateTime.UtcNow:yyyy-MM-dd HH:mm}"
                : name.Trim(),
            CreatedAtUtc = DateTime.UtcNow,
            UpdatedAtUtc = DateTime.UtcNow
        };

        db.Plans.Add(plan);
        await db.SaveChangesAsync();

        // сохраняем содержимое как копию, привязывая к plan.Id
        var clonedTables = CloneTablesForPlan(plan.Id, tables);
        var clonedSummary = CloneSummaryForPlan(plan.Id, summary);

        db.Tables.AddRange(clonedTables);
        db.SummaryTables.Add(clonedSummary);

        await db.SaveChangesAsync();

        return plan.Id;
    }

    public async Task SavePlanAsync(int planId, string? name, List<PlanTable> tables, SummaryTable summary)
    {
        await using var db = CreateContext();

        var plan = await db.Plans.FirstAsync(p => p.Id == planId);

        if (!string.IsNullOrWhiteSpace(name))
            plan.Name = name.Trim();

        plan.UpdatedAtUtc = DateTime.UtcNow;

        // Удаляем старые данные плана (проще и надёжнее на текущем этапе)
        var oldTables = await db.Tables
            .Where(t => t.PlanId == planId)
            .Include(t => t.Rows)
            .ToListAsync();

        db.Tables.RemoveRange(oldTables);

        var oldSummary = await db.SummaryTables
            .Where(s => s.PlanId == planId)
            .Include(s => s.Rows)
            .FirstOrDefaultAsync();

        if (oldSummary != null)
            db.SummaryTables.Remove(oldSummary);

        await db.SaveChangesAsync();

        // Добавляем новые данные
        var clonedTables = CloneTablesForPlan(planId, tables);
        var clonedSummary = CloneSummaryForPlan(planId, summary);

        db.Tables.AddRange(clonedTables);
        db.SummaryTables.Add(clonedSummary);

        await db.SaveChangesAsync();
    }

    public async Task<(List<PlanTable> Tables, SummaryTable? Summary)> LoadPlanAsync(int planId)
    {
        await using var db = CreateContext();

        var tables = await db.Tables
            .Where(t => t.PlanId == planId)
            .Include(t => t.Rows.OrderBy(r => r.RowOrder))
            .OrderBy(t => t.SheetName)
            .ToListAsync();

        var summary = await db.SummaryTables
            .Where(s => s.PlanId == planId)
            .Include(s => s.Rows.OrderBy(r => r.RowOrder))
            .FirstOrDefaultAsync();

        return (tables, summary);
    }

    public async Task DeletePlanAsync(int planId)
    {
        await using var db = CreateContext();
        var plan = await db.Plans.FirstAsync(p => p.Id == planId);
        db.Plans.Remove(plan); // каскадно удалит таблицы/строки/итоги
        await db.SaveChangesAsync();
    }

    // -----------------------------
    //  C) Клонирование (чтобы EF не путался в tracking)
    // -----------------------------

    private static List<PlanTable> CloneTablesForPlan(int planId, List<PlanTable> tables)
    {
        var result = new List<PlanTable>(tables.Count);

        foreach (var t in tables)
        {
            var nt = new PlanTable
            {
                PlanId = planId,

                SheetName = t.SheetName,
                SemesterTitle = t.SemesterTitle,

                Sem1Plan = t.Sem1Plan,
                Sem1Fact = t.Sem1Fact,
                Sem2Plan = t.Sem2Plan,
                Sem2Fact = t.Sem2Fact,
                YearPlan = t.YearPlan,
                YearFact = t.YearFact,

                Rows = new List<PlanRow>()
            };

            foreach (var r in t.Rows.OrderBy(x => x.RowOrder))
            {
                var nr = new PlanRow
                {
                    RowOrder = r.RowOrder,
                    IsSummary = r.IsSummary,

                    Number = r.Number,
                    DisciplineName = r.DisciplineName,
                    FacultyGroup = r.FacultyGroup,

                    Course = r.Course,
                    Streams = r.Streams,
                    Groups = r.Groups,
                    Students = r.Students,

                    Lek = r.Lek,
                    Pr = r.Pr,
                    Lab = r.Lab,
                    Ksr = r.Ksr,
                    Kp = r.Kp,
                    Kr = r.Kr,
                    KontrolRab = r.KontrolRab,
                    Zach = r.Zach,
                    DifZach = r.DifZach,
                    Exz = r.Exz,
                    GosExz = r.GosExz,
                    Gek = r.Gek,
                    RukVkr = r.RukVkr,
                    Rec = r.Rec,
                    UchPr = r.UchPr,
                    PrPr = r.PrPr,
                    PredPr = r.PredPr,

                    Total = r.Total,
                    Note = r.Note,

                    PlanTable = nt
                };

                nt.Rows.Add(nr);
            }

            result.Add(nt);
        }

        return result;
    }

    private static SummaryTable CloneSummaryForPlan(int planId, SummaryTable summary)
    {
        var ns = new SummaryTable
        {
            PlanId = planId,
            CreatedAtUtc = DateTime.UtcNow,
            Rows = new List<SummaryRow>()
        };

        foreach (var r in summary.Rows.OrderBy(x => x.RowOrder))
        {
            ns.Rows.Add(new SummaryRow
            {
                RowOrder = r.RowOrder,
                Code = r.Code,
                WorkName = r.WorkName,
                Sem1Plan = r.Sem1Plan,
                Sem1Fact = r.Sem1Fact,
                Sem2Plan = r.Sem2Plan,
                Sem2Fact = r.Sem2Fact,
                YearPlan = r.YearPlan,
                YearFact = r.YearFact,
                IsTotalRow = r.IsTotalRow,
                SummaryTable = ns
            });
        }

        return ns;
    }
}