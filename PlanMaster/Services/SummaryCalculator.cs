using PlanMaster.Models;

namespace PlanMaster.Services;

public class SummaryCalculator
{
    // ВАЖНО: чтобы совпало с твоим примером (394/401/795),
    // считаем "Учебная работа" только из Table 2 и Table 3.
    private static readonly HashSet<string> TeachingSheets = new(StringComparer.OrdinalIgnoreCase)
    {
        "Table 2", "Table 3"
    };

    public SummaryTable CreateOrUpdate(SummaryTable? existing, IReadOnlyList<PlanTable> importedTables)
    {
        var table = existing ?? new SummaryTable();

        if (table.Rows.Count == 0)
            table.Rows = CreateDefaultRows();

        // 1) Авто: Учебная работа
        var teaching = SumTotals(importedTables.Where(t => TeachingSheets.Contains(t.SheetName)));

        var row1 = table.Rows.First(r => r.Code == "1");
        row1.Sem1Plan = teaching.Sem1Plan;
        row1.Sem1Fact = teaching.Sem1Fact;
        row1.Sem2Plan = teaching.Sem2Plan;
        row1.Sem2Fact = teaching.Sem2Fact;
        row1.YearPlan = teaching.YearPlan;
        row1.YearFact = teaching.YearFact;

        // 2) Пересчёт строки ИТОГО
        RecalcTotalRow(table);

        return table;
    }

    private static (int Sem1Plan, int Sem1Fact, int Sem2Plan, int Sem2Fact, int YearPlan, int YearFact)
        SumTotals(IEnumerable<PlanTable> tables)
    {
        int s1p = 0, s1f = 0, s2p = 0, s2f = 0, yp = 0, yf = 0;
        foreach (var t in tables)
        {
            s1p += t.Sem1Plan ?? 0;
            s1f += t.Sem1Fact ?? 0;
            s2p += t.Sem2Plan ?? 0;
            s2f += t.Sem2Fact ?? 0;
            yp += t.YearPlan ?? 0;
            yf += t.YearFact ?? 0;
        }
        return (s1p, s1f, s2p, s2f, yp, yf);
    }

    private static void RecalcTotalRow(SummaryTable table)
    {
        var total = table.Rows.First(r => r.IsTotalRow);

        var rows = table.Rows.Where(r => !r.IsTotalRow).ToList();

        total.Sem1Plan = rows.Sum(r => r.Sem1Plan);
        total.Sem1Fact = rows.Sum(r => r.Sem1Fact);
        total.Sem2Plan = rows.Sum(r => r.Sem2Plan);
        total.Sem2Fact = rows.Sum(r => r.Sem2Fact);
        total.YearPlan = rows.Sum(r => r.YearPlan);
        total.YearFact = rows.Sum(r => r.YearFact);
    }

    private static List<SummaryRow> CreateDefaultRows()
    {
        // Структура как на твоём скрине
        var rows = new List<SummaryRow>
        {
            new() { RowOrder = 1, Code="1", WorkName="Учебная работа" },

            new() { RowOrder = 2, Code="2", WorkName="Учебно-методическая работа" },
            new() { RowOrder = 3, Code="2.1", WorkName="Разработка методического обеспечения учебного процесса" },
            new() { RowOrder = 4, Code="2.2", WorkName="Подготовка к изданию учебно-методических и научных разработок" },
            new() { RowOrder = 5, Code="2.3", WorkName="Совершенствование учебно-материальной базы" },

            new() { RowOrder = 6, Code="3", WorkName="Организационно-методическая работа" },

            new() { RowOrder = 7, Code="4", WorkName="Организация научно-исследовательской работы студентов" },

            new() { RowOrder = 8, Code="5", WorkName="Научно-исследовательская работа" },

            new() { RowOrder = 9, Code="6", WorkName="Повышение квалификации" },

            new() { RowOrder = 10, Code="7", WorkName="Внеучебная работа" },

            new() { RowOrder = 11, Code="8", WorkName="Другие виды работ" },

            new() { RowOrder = 12, Code="", WorkName="Итого:", IsTotalRow = true }
        };

        // Значения по умолчанию = 0, пользователь редактирует остальные строки вручную
        return rows;
    }
}