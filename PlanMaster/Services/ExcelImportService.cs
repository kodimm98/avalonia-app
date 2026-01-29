using ClosedXML.Excel;
using PlanMaster.Models;

namespace PlanMaster.Services;

public class ExcelImportService
{
    private const int ColA = 1;   // №
    private const int ColB = 2;   // Наименование
    private const int ColC = 3;   // Факультет...
    private const int ColTotal = 25; // Y = "ВСЕГО"
    private const int ColNote = 26;  // Z = "Примечание"

    public List<PlanTable> Import(string filePath)
    {
        using var wb = new XLWorkbook(filePath);
        var result = new List<PlanTable>();

        foreach (var ws in wb.Worksheets)
        {
            if (!ws.Name.StartsWith("Table", StringComparison.OrdinalIgnoreCase))
                continue;

            var lastRow = ws.LastRowUsed()?.RowNumber() ?? 0;
            if (lastRow < 3) continue;

            var headerRow = FindHeaderRow(ws, lastRow);
            if (headerRow == null) continue;

            var semRow = FindSemesterRow(ws, lastRow);

            // Если семестр найден — данные с следующей строки.
            // Если не найден — обычно 2 строки заголовков.
            int dataStart = semRow != null ? semRow.Value + 1 : headerRow.Value + 2;

            var table = new PlanTable
            {
                SheetName = ws.Name,
                SemesterTitle = semRow != null ? ws.Cell(semRow.Value, ColB).GetString().Trim() : null,
            };

            int rowOrder = 0;

            int? lastNumber = null;
            string lastDiscipline = "";

            string? lastTotalHeader = null;

            for (int r = dataStart; r <= lastRow; r++)
            {
                var aStr = ws.Cell(r, ColA).GetString().Trim();
                var b = ws.Cell(r, ColB).GetString().Trim();
                var c = ws.Cell(r, ColC).GetString().Trim();

                // Пропуск пустых
                if (IsRowEmpty(ws, r))
                    continue;

                // --- ИТОГИ: "Итого за ..." ---
                if (IsTotalHeader(b))
                {
                    lastTotalHeader = b;

                    // если "Поручено/Выполнено" на этой же строке
                    if (IsPlanOrFact(c))
                    {
                        var summaryRow = CreateSummaryRow(ws, r, rowOrder++, headerB: b, kindC: c);
                        table.Rows.Add(summaryRow);

                        ApplyTotals(table, b, c, summaryRow.Total);
                    }

                    continue;
                }

                // продолжение итогов: B пустая, C="Поручено/Выполнено", заголовок был строкой выше
                if (string.IsNullOrWhiteSpace(b) && IsPlanOrFact(c) && !string.IsNullOrWhiteSpace(lastTotalHeader))
                {
                    var summaryRow = CreateSummaryRow(ws, r, rowOrder++, headerB: lastTotalHeader!, kindC: c);
                    table.Rows.Add(summaryRow);

                    ApplyTotals(table, lastTotalHeader!, c, summaryRow.Total);
                    continue;
                }

                // если вдруг C="Поручено/Выполнено" без заголовка — не добавляем как дисциплину
                if (IsPlanOrFact(c) && string.IsNullOrWhiteSpace(b))
                    continue;

                // --- ОБЫЧНЫЕ СТРОКИ ДИСЦИПЛИН ---
                var discipline = b;

                // forward-fill для merged cells
                var number = ParseNullableInt(aStr) ?? lastNumber;

                if (!string.IsNullOrWhiteSpace(discipline))
                    lastDiscipline = discipline;

                if (number == null && string.IsNullOrWhiteSpace(lastDiscipline))
                    continue;

                lastNumber = number;

                var row = new PlanRow
                {
                    RowOrder = rowOrder++,
                    IsSummary = false,

                    Number = number,
                    DisciplineName = lastDiscipline,
                    FacultyGroup = c,

                    Course = GetInt(ws.Cell(r, 4)),
                    Streams = GetInt(ws.Cell(r, 5)),
                    Groups = GetInt(ws.Cell(r, 6)),
                    Students = GetInt(ws.Cell(r, 7)),

                    Lek = GetInt(ws.Cell(r, 8)),
                    Pr = GetInt(ws.Cell(r, 9)),
                    Lab = GetInt(ws.Cell(r, 10)),
                    Ksr = GetInt(ws.Cell(r, 11)),
                    Kp = GetInt(ws.Cell(r, 12)),
                    Kr = GetInt(ws.Cell(r, 13)),
                    KontrolRab = GetInt(ws.Cell(r, 14)),
                    Zach = GetInt(ws.Cell(r, 15)),
                    DifZach = GetInt(ws.Cell(r, 16)),
                    Exz = GetInt(ws.Cell(r, 17)),
                    GosExz = GetInt(ws.Cell(r, 18)),
                    Gek = GetInt(ws.Cell(r, 19)),
                    RukVkr = GetInt(ws.Cell(r, 20)),
                    Rec = GetInt(ws.Cell(r, 21)),
                    UchPr = GetInt(ws.Cell(r, 22)),
                    PrPr = GetInt(ws.Cell(r, 23)),
                    PredPr = GetInt(ws.Cell(r, 24)),

                    Total = GetInt(ws.Cell(r, ColTotal)),
                    Note = ws.Cell(r, ColNote).GetString().Trim()
                };

                table.Rows.Add(row);
            }

            // добавляем лист если есть строки или итоги
            if (table.Rows.Count > 0 ||
                table.Sem1Plan.HasValue || table.Sem2Plan.HasValue || table.YearPlan.HasValue ||
                table.Sem1Fact.HasValue || table.Sem2Fact.HasValue || table.YearFact.HasValue)
            {
                result.Add(table);
            }
        }

        return result;
    }

    private static PlanRow CreateSummaryRow(IXLWorksheet ws, int r, int rowOrder, string headerB, string kindC)
    {
        return new PlanRow
        {
            RowOrder = rowOrder,
            IsSummary = true,

            Number = null,
            DisciplineName = headerB, // "Итого за 1 семестр:"
            FacultyGroup = kindC,     // "Поручено" / "Выполнено"

            Course = GetInt(ws.Cell(r, 4)),
            Streams = GetInt(ws.Cell(r, 5)),
            Groups = GetInt(ws.Cell(r, 6)),
            Students = GetInt(ws.Cell(r, 7)),

            Lek = GetInt(ws.Cell(r, 8)),
            Pr = GetInt(ws.Cell(r, 9)),
            Lab = GetInt(ws.Cell(r, 10)),
            Ksr = GetInt(ws.Cell(r, 11)),
            Kp = GetInt(ws.Cell(r, 12)),
            Kr = GetInt(ws.Cell(r, 13)),
            KontrolRab = GetInt(ws.Cell(r, 14)),
            Zach = GetInt(ws.Cell(r, 15)),
            DifZach = GetInt(ws.Cell(r, 16)),
            Exz = GetInt(ws.Cell(r, 17)),
            GosExz = GetInt(ws.Cell(r, 18)),
            Gek = GetInt(ws.Cell(r, 19)),
            RukVkr = GetInt(ws.Cell(r, 20)),
            Rec = GetInt(ws.Cell(r, 21)),
            UchPr = GetInt(ws.Cell(r, 22)),
            PrPr = GetInt(ws.Cell(r, 23)),
            PredPr = GetInt(ws.Cell(r, 24)),

            Total = GetInt(ws.Cell(r, ColTotal)),
            Note = ws.Cell(r, ColNote).GetString().Trim()
        };
    }

    private static int? FindHeaderRow(IXLWorksheet ws, int lastRow)
    {
        for (int r = 1; r <= Math.Min(lastRow, 40); r++)
        {
            var a = ws.Cell(r, ColA).GetString().Trim();
            if (a.Equals("№ п.п.", StringComparison.OrdinalIgnoreCase))
                return r;
        }
        return null;
    }

    private static int? FindSemesterRow(IXLWorksheet ws, int lastRow)
    {
        for (int r = 1; r <= Math.Min(lastRow, 80); r++)
        {
            var b = ws.Cell(r, ColB).GetString().Trim();
            if (b.Contains("семестр", StringComparison.OrdinalIgnoreCase))
                return r;
        }
        return null;
    }

    private static bool IsTotalHeader(string b)
        => b.StartsWith("Итого за 1 семестр", StringComparison.OrdinalIgnoreCase)
        || b.StartsWith("Итого за 2 семестр", StringComparison.OrdinalIgnoreCase)
        || b.StartsWith("Итого за год", StringComparison.OrdinalIgnoreCase);

    private static bool IsPlanOrFact(string c)
        => c.Equals("Поручено", StringComparison.OrdinalIgnoreCase)
        || c.Equals("Выполнено", StringComparison.OrdinalIgnoreCase);

    private static void ApplyTotals(PlanTable t, string headerB, string kindC, int? total)
    {
        if (total == null) return;

        bool isPlan = kindC.Equals("Поручено", StringComparison.OrdinalIgnoreCase);
        bool isFact = kindC.Equals("Выполнено", StringComparison.OrdinalIgnoreCase);

        if (!isPlan && !isFact) return;

        if (headerB.StartsWith("Итого за 1 семестр", StringComparison.OrdinalIgnoreCase))
        {
            if (isPlan) t.Sem1Plan = total;
            if (isFact) t.Sem1Fact = total;
        }
        else if (headerB.StartsWith("Итого за 2 семестр", StringComparison.OrdinalIgnoreCase))
        {
            if (isPlan) t.Sem2Plan = total;
            if (isFact) t.Sem2Fact = total;
        }
        else if (headerB.StartsWith("Итого за год", StringComparison.OrdinalIgnoreCase))
        {
            if (isPlan) t.YearPlan = total;
            if (isFact) t.YearFact = total;
        }
    }

    private static bool IsRowEmpty(IXLWorksheet ws, int r)
    {
        // A..Z
        for (int col = 1; col <= ColNote; col++)
        {
            var cell = ws.Cell(r, col);

            if (!cell.IsEmpty())
                return false;

            if (!string.IsNullOrWhiteSpace(cell.GetString()))
                return false;
        }
        return true;
    }

    private static int? GetInt(IXLCell cell)
    {
        if (cell.IsEmpty()) return null;

        if (cell.TryGetValue<double>(out var d))
            return (int)Math.Round(d);

        var s = cell.GetString().Trim();
        return ParseNullableInt(s);
    }

    private static int? ParseNullableInt(string? s)
    {
        if (string.IsNullOrWhiteSpace(s)) return null;
        return int.TryParse(s, out var x) ? x : null;
    }
}