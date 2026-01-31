using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using PlanMaster.Models;

namespace PlanMaster.ViewModels;

public partial class MainWindowViewModel
{
    public bool CanDeleteMethodicalProcessRow => SelectedMethodicalProcessRow != null;
    public bool CanDeleteMethodicalPublishingRow => SelectedMethodicalPublishingRow != null;
    public bool CanDeleteMethodicalBaseRow => SelectedMethodicalBaseRow != null;
    public bool CanDeleteOrganizationalResearchRow => SelectedOrganizationalResearchRow != null;
    public bool CanDeleteResearchWorkRow => SelectedResearchWorkRow != null;
    public bool CanDeleteQualificationRow => SelectedQualificationRow != null;
    public bool CanDeleteExtracurricularRow => SelectedExtracurricularRow != null;
    public bool CanDeleteOtherWorkRow => SelectedOtherWorkRow != null;

    public void AddMethodicalProcessRow()
        => MethodicalProcessRows.Add(new MethodWorkRow { Category = MethodicalProcessCategory });

    public void AddMethodicalPublishingRow()
        => MethodicalPublishingRows.Add(new MethodWorkRow { Category = MethodicalPublishingCategory });

    public void AddMethodicalBaseRow()
        => MethodicalBaseRows.Add(new MethodWorkRow { Category = MethodicalBaseCategory });

    public void AddOrganizationalResearchRow()
        => OrganizationalResearchRows.Add(new MethodWorkRow { Category = OrganizationalResearchCategory });

    public void AddResearchWorkRow()
        => ResearchWorkRows.Add(new MethodWorkRow { Category = ResearchWorkCategory });

    public void AddQualificationRow()
        => QualificationRows.Add(new MethodWorkRow { Category = QualificationCategory });

    public void AddExtracurricularRow()
        => ExtracurricularRows.Add(new MethodWorkRow { Category = ExtracurricularCategory });

    public void AddOtherWorkRow()
        => OtherWorkRows.Add(new MethodWorkRow { Category = OtherWorkCategory });

    public void DeleteMethodicalProcessRow()
    {
        if (SelectedMethodicalProcessRow == null) return;
        MethodicalProcessRows.Remove(SelectedMethodicalProcessRow);
        SelectedMethodicalProcessRow = null;
        RefreshSummaryRows(BuildRecalcedSummaryFromUi());
    }

    public void DeleteMethodicalPublishingRow()
    {
        if (SelectedMethodicalPublishingRow == null) return;
        MethodicalPublishingRows.Remove(SelectedMethodicalPublishingRow);
        SelectedMethodicalPublishingRow = null;
        RefreshSummaryRows(BuildRecalcedSummaryFromUi());
    }

    public void DeleteMethodicalBaseRow()
    {
        if (SelectedMethodicalBaseRow == null) return;
        MethodicalBaseRows.Remove(SelectedMethodicalBaseRow);
        SelectedMethodicalBaseRow = null;
        RefreshSummaryRows(BuildRecalcedSummaryFromUi());
    }

    public void DeleteOrganizationalResearchRow()
    {
        if (SelectedOrganizationalResearchRow == null) return;
        OrganizationalResearchRows.Remove(SelectedOrganizationalResearchRow);
        SelectedOrganizationalResearchRow = null;
    }

    public void DeleteResearchWorkRow()
    {
        if (SelectedResearchWorkRow == null) return;
        ResearchWorkRows.Remove(SelectedResearchWorkRow);
        SelectedResearchWorkRow = null;
    }

    public void DeleteQualificationRow()
    {
        if (SelectedQualificationRow == null) return;
        QualificationRows.Remove(SelectedQualificationRow);
        SelectedQualificationRow = null;
    }

    public void DeleteExtracurricularRow()
    {
        if (SelectedExtracurricularRow == null) return;
        ExtracurricularRows.Remove(SelectedExtracurricularRow);
        SelectedExtracurricularRow = null;
    }

    public void DeleteOtherWorkRow()
    {
        if (SelectedOtherWorkRow == null) return;
        OtherWorkRows.Remove(SelectedOtherWorkRow);
        SelectedOtherWorkRow = null;
    }

    partial void OnSelectedMethodicalProcessRowChanged(MethodWorkRow? value)
        => OnPropertyChanged(nameof(CanDeleteMethodicalProcessRow));

    partial void OnSelectedMethodicalPublishingRowChanged(MethodWorkRow? value)
        => OnPropertyChanged(nameof(CanDeleteMethodicalPublishingRow));

    partial void OnSelectedMethodicalBaseRowChanged(MethodWorkRow? value)
        => OnPropertyChanged(nameof(CanDeleteMethodicalBaseRow));

    partial void OnSelectedOrganizationalResearchRowChanged(MethodWorkRow? value)
        => OnPropertyChanged(nameof(CanDeleteOrganizationalResearchRow));

    partial void OnSelectedResearchWorkRowChanged(MethodWorkRow? value)
        => OnPropertyChanged(nameof(CanDeleteResearchWorkRow));

    partial void OnSelectedQualificationRowChanged(MethodWorkRow? value)
        => OnPropertyChanged(nameof(CanDeleteQualificationRow));

    partial void OnSelectedExtracurricularRowChanged(MethodWorkRow? value)
        => OnPropertyChanged(nameof(CanDeleteExtracurricularRow));

    partial void OnSelectedOtherWorkRowChanged(MethodWorkRow? value)
        => OnPropertyChanged(nameof(CanDeleteOtherWorkRow));

    private void ResetMethodicalRows()
    {
        MethodicalProcessRows.Clear();
        MethodicalPublishingRows.Clear();
        MethodicalBaseRows.Clear();
        OrganizationalResearchRows.Clear();
        ResearchWorkRows.Clear();
        QualificationRows.Clear();
        ExtracurricularRows.Clear();
        OtherWorkRows.Clear();

        MethodicalProcessRows.Add(new MethodWorkRow { Category = MethodicalProcessCategory });
        MethodicalPublishingRows.Add(new MethodWorkRow { Category = MethodicalPublishingCategory });
        MethodicalBaseRows.Add(new MethodWorkRow { Category = MethodicalBaseCategory });
        OrganizationalResearchRows.Add(new MethodWorkRow { Category = OrganizationalResearchCategory });
        ResearchWorkRows.Add(new MethodWorkRow { Category = ResearchWorkCategory });
        QualificationRows.Add(new MethodWorkRow { Category = QualificationCategory });
        ExtracurricularRows.Add(new MethodWorkRow { Category = ExtracurricularCategory });
        OtherWorkRows.Add(new MethodWorkRow { Category = OtherWorkCategory });
    }

    private void EnsureMethodicalRows()
    {
        if (MethodicalProcessRows.Count == 0
            && MethodicalPublishingRows.Count == 0
            && MethodicalBaseRows.Count == 0
            && OrganizationalResearchRows.Count == 0
            && ResearchWorkRows.Count == 0
            && QualificationRows.Count == 0
            && ExtracurricularRows.Count == 0
            && OtherWorkRows.Count == 0)
        {
            ResetMethodicalRows();
        }
    }

    private void LoadMethodical(MethodWorkTable? table)
    {
        MethodicalProcessRows.Clear();
        MethodicalPublishingRows.Clear();
        MethodicalBaseRows.Clear();
        OrganizationalResearchRows.Clear();
        ResearchWorkRows.Clear();
        QualificationRows.Clear();
        ExtracurricularRows.Clear();
        OtherWorkRows.Clear();

        if (table == null)
        {
            ResetMethodicalRows();
            return;
        }

        foreach (var row in table.Rows.OrderBy(r => r.RowOrder))
        {
            var target = row.Category switch
            {
                MethodicalProcessCategory => MethodicalProcessRows,
                MethodicalPublishingCategory => MethodicalPublishingRows,
                MethodicalBaseCategory => MethodicalBaseRows,
                OrganizationalResearchCategory => OrganizationalResearchRows,
                ResearchWorkCategory => ResearchWorkRows,
                QualificationCategory => QualificationRows,
                ExtracurricularCategory => ExtracurricularRows,
                OtherWorkCategory => OtherWorkRows,
                _ => MethodicalProcessRows
            };

            target.Add(new MethodWorkRow
            {
                Category = row.Category,
                WorkName = row.WorkName,
                TimeHours = row.TimeHours,
                Deadline = row.Deadline,
                CompletionNote = row.CompletionNote
            });
        }

        if (MethodicalProcessRows.Count == 0)
            MethodicalProcessRows.Add(new MethodWorkRow { Category = MethodicalProcessCategory });
        if (MethodicalPublishingRows.Count == 0)
            MethodicalPublishingRows.Add(new MethodWorkRow { Category = MethodicalPublishingCategory });
        if (MethodicalBaseRows.Count == 0)
            MethodicalBaseRows.Add(new MethodWorkRow { Category = MethodicalBaseCategory });
        if (OrganizationalResearchRows.Count == 0)
            OrganizationalResearchRows.Add(new MethodWorkRow { Category = OrganizationalResearchCategory });
        if (ResearchWorkRows.Count == 0)
            ResearchWorkRows.Add(new MethodWorkRow { Category = ResearchWorkCategory });
        if (QualificationRows.Count == 0)
            QualificationRows.Add(new MethodWorkRow { Category = QualificationCategory });
        if (ExtracurricularRows.Count == 0)
            ExtracurricularRows.Add(new MethodWorkRow { Category = ExtracurricularCategory });
        if (OtherWorkRows.Count == 0)
            OtherWorkRows.Add(new MethodWorkRow { Category = OtherWorkCategory });
    }

    private MethodWorkTable BuildMethodicalFromUi()
    {
        var table = new MethodWorkTable
        {
            Rows = new List<MethodWorkRow>()
        };

        var rowOrder = 0;
        AddRows(MethodicalProcessRows, MethodicalProcessCategory);
        AddRows(MethodicalPublishingRows, MethodicalPublishingCategory);
        AddRows(MethodicalBaseRows, MethodicalBaseCategory);
        AddRows(OrganizationalResearchRows, OrganizationalResearchCategory);
        AddRows(ResearchWorkRows, ResearchWorkCategory);
        AddRows(QualificationRows, QualificationCategory);
        AddRows(ExtracurricularRows, ExtracurricularCategory);
        AddRows(OtherWorkRows, OtherWorkCategory);

        return table;

        void AddRows(IEnumerable<MethodWorkRow> rows, string category)
        {
            foreach (var r in rows)
            {
                if (string.IsNullOrWhiteSpace(r.WorkName)
                    && r.TimeHours is null
                    && string.IsNullOrWhiteSpace(r.Deadline)
                    && string.IsNullOrWhiteSpace(r.CompletionNote))
                {
                    continue;
                }

                table.Rows.Add(new MethodWorkRow
                {
                    RowOrder = rowOrder++,
                    Category = category,
                    WorkName = r.WorkName ?? "",
                    TimeHours = r.TimeHours,
                    Deadline = r.Deadline ?? "",
                    CompletionNote = r.CompletionNote ?? ""
                });
            }
        }
    }

    private SummaryTable BuildRecalcedSummaryFromUi()
    {
        var summary = BuildSummaryFromUi();
        ApplyMethodicalSummary(summary, BuildMethodicalFromUi());
        var tablesToUse = Tables.Where(t => t.IncludeInSummary).ToList();
        return _summaryCalculator.CreateOrUpdate(summary, tablesToUse);
    }

    public void RecalculateSummaryFromTables()
    {
        var recalced = BuildRecalcedSummaryFromUi();
        RefreshSummaryRows(recalced);
    }

    public void RecalculateTableTotals(PlanTable? table)
    {
        if (table == null) return;
        RecalculateTeachingTableTotals(table);
        RecalculateTotalsFromSummaryRows(table);
        RecalculateSummaryFromTables();
    }

    private static void RecalculateTeachingTableTotals(PlanTable table)
    {
        foreach (var row in table.Rows.Where(r => !r.IsSummary))
        {
            row.Total = SumHours(row);
        }

        var total = table.Rows.Where(r => !r.IsSummary).Sum(r => r.Total ?? 0);
        foreach (var summaryRow in table.Rows.Where(r => r.IsSummary))
        {
            summaryRow.Total = total;
        }
    }

    private static void RecalculateTotalsFromSummaryRows(PlanTable table)
    {
        var summaryRows = table.Rows.Where(r => r.IsSummary).ToList();

        table.Sem1Plan = FindTotal(summaryRows, "Итого за 1 семестр", "Поручено");
        table.Sem1Fact = FindTotal(summaryRows, "Итого за 1 семестр", "Выполнено");
        table.Sem2Plan = FindTotal(summaryRows, "Итого за 2 семестр", "Поручено");
        table.Sem2Fact = FindTotal(summaryRows, "Итого за 2 семестр", "Выполнено");
        table.YearPlan = FindTotal(summaryRows, "Итого за год", "Поручено");
        table.YearFact = FindTotal(summaryRows, "Итого за год", "Выполнено");
    }

    private static int? FindTotal(IEnumerable<PlanRow> rows, string header, string kind)
        => rows.FirstOrDefault(r => r.DisciplineName.StartsWith(header, StringComparison.OrdinalIgnoreCase)
                                    && r.FacultyGroup.Equals(kind, StringComparison.OrdinalIgnoreCase))
                ?.Total;

    private static int SumHours(PlanRow row)
        => (row.Lek ?? 0)
           + (row.Pr ?? 0)
           + (row.Lab ?? 0)
           + (row.Ksr ?? 0)
           + (row.Kp ?? 0)
           + (row.Kr ?? 0)
           + (row.KontrolRab ?? 0)
           + (row.Zach ?? 0)
           + (row.DifZach ?? 0)
           + (row.Exz ?? 0)
           + (row.GosExz ?? 0)
           + (row.Gek ?? 0)
           + (row.RukVkr ?? 0)
           + (row.Rec ?? 0)
           + (row.UchPr ?? 0)
           + (row.PrPr ?? 0)
           + (row.PredPr ?? 0);

    private void RefreshSummaryRows(SummaryTable summary)
    {
        SummaryRows.Clear();
        foreach (var r in summary.Rows.OrderBy(r => r.RowOrder))
            SummaryRows.Add(r);
    }

    private void ApplyMethodicalSummary(SummaryTable summary, MethodWorkTable methodical)
    {
        if (summary.Rows.Count == 0) return;

        var row21 = summary.Rows.FirstOrDefault(r => r.Code == "2.1");
        var row22 = summary.Rows.FirstOrDefault(r => r.Code == "2.2");
        var row23 = summary.Rows.FirstOrDefault(r => r.Code == "2.3");

        if (row21 == null || row22 == null || row23 == null)
            return;

        var totals = new Dictionary<string, (int Sem1, int Sem2)>
        {
            [MethodicalProcessCategory] = (0, 0),
            [MethodicalPublishingCategory] = (0, 0),
            [MethodicalBaseCategory] = (0, 0),
            [OrganizationalResearchCategory] = (0, 0),
            [ResearchWorkCategory] = (0, 0),
            [QualificationCategory] = (0, 0),
            [ExtracurricularCategory] = (0, 0),
            [OtherWorkCategory] = (0, 0)
        };

        foreach (var row in methodical.Rows)
        {
            if (!totals.ContainsKey(row.Category))
                continue;

            if (row.TimeHours is null || row.TimeHours.Value == 0)
                continue;

            if (!TryParseDeadline(row.Deadline, out var date))
                continue;

            var isFirst = IsFirstSemester(date);
            var current = totals[row.Category];
            if (isFirst)
                totals[row.Category] = (current.Sem1 + row.TimeHours.Value, current.Sem2);
            else
                totals[row.Category] = (current.Sem1, current.Sem2 + row.TimeHours.Value);
        }

        SetMethodicalRow(row21, totals[MethodicalProcessCategory]);
        SetMethodicalRow(row22, totals[MethodicalPublishingCategory]);
        SetMethodicalRow(row23, totals[MethodicalBaseCategory]);

        SetSummaryRow(summary, "4", totals[OrganizationalResearchCategory]);
        SetSummaryRow(summary, "5", totals[ResearchWorkCategory]);
        SetSummaryRow(summary, "6", totals[QualificationCategory]);
        SetSummaryRow(summary, "7", totals[ExtracurricularCategory]);
        SetSummaryRow(summary, "8", totals[OtherWorkCategory]);
    }

    private static void SetMethodicalRow(SummaryRow row, (int Sem1, int Sem2) totals)
    {
        row.Sem1Plan = totals.Sem1;
        row.Sem2Plan = totals.Sem2;
        row.YearPlan = totals.Sem1 + totals.Sem2;
    }

    private static void SetSummaryRow(SummaryTable summary, string code, (int Sem1, int Sem2) totals)
    {
        var row = summary.Rows.FirstOrDefault(r => r.Code == code);
        if (row == null) return;
        row.Sem1Plan = totals.Sem1;
        row.Sem2Plan = totals.Sem2;
        row.YearPlan = totals.Sem1 + totals.Sem2;
    }

    private static bool TryParseDeadline(string deadline, out DateTime date)
    {
        if (string.IsNullOrWhiteSpace(deadline))
        {
            date = default;
            return false;
        }

        var culture = CultureInfo.GetCultureInfo("ru-RU");
        return DateTime.TryParse(deadline, culture, DateTimeStyles.None, out date)
            || DateTime.TryParse(deadline, CultureInfo.InvariantCulture, DateTimeStyles.None, out date);
    }

    private static bool IsFirstSemester(DateTime date)
        => date.Month >= 9 || date.Month <= 2;
}
