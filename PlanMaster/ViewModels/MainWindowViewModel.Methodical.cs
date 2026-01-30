using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using PlanMaster.Models;

namespace PlanMaster.ViewModels;

public partial class MainWindowViewModel
{
    public void AddMethodicalProcessRow()
        => MethodicalProcessRows.Add(new MethodWorkRow { Category = MethodicalProcessCategory });

    public void AddMethodicalPublishingRow()
        => MethodicalPublishingRows.Add(new MethodWorkRow { Category = MethodicalPublishingCategory });

    public void AddMethodicalBaseRow()
        => MethodicalBaseRows.Add(new MethodWorkRow { Category = MethodicalBaseCategory });

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

    partial void OnSelectedMethodicalProcessRowChanged(MethodWorkRow? value)
        => OnPropertyChanged(nameof(CanDeleteMethodicalProcessRow));

    partial void OnSelectedMethodicalPublishingRowChanged(MethodWorkRow? value)
        => OnPropertyChanged(nameof(CanDeleteMethodicalPublishingRow));

    partial void OnSelectedMethodicalBaseRowChanged(MethodWorkRow? value)
        => OnPropertyChanged(nameof(CanDeleteMethodicalBaseRow));

    private void ResetMethodicalRows()
    {
        MethodicalProcessRows.Clear();
        MethodicalPublishingRows.Clear();
        MethodicalBaseRows.Clear();

        MethodicalProcessRows.Add(new MethodWorkRow { Category = MethodicalProcessCategory });
        MethodicalPublishingRows.Add(new MethodWorkRow { Category = MethodicalPublishingCategory });
        MethodicalBaseRows.Add(new MethodWorkRow { Category = MethodicalBaseCategory });
    }

    private void LoadMethodical(MethodWorkTable? table)
    {
        MethodicalProcessRows.Clear();
        MethodicalPublishingRows.Clear();
        MethodicalBaseRows.Clear();

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
        return _summaryCalculator.CreateOrUpdate(summary, Tables.ToList());
    }

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
            [MethodicalBaseCategory] = (0, 0)
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
    }

    private static void SetMethodicalRow(SummaryRow row, (int Sem1, int Sem2) totals)
    {
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
