using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using CommunityToolkit.Mvvm.ComponentModel;
using PlanMaster.Models;
using PlanMaster.Services;
using PlanMaster.Data;

namespace PlanMaster.ViewModels;

public partial class MainWindowViewModel : ViewModelBase
{
    public const string MethodicalProcessCategory = "Разработка методического обеспечения учебного процесса";
    public const string MethodicalPublishingCategory = "Подготовка к изданию учебно-методических разработок";
    public const string MethodicalBaseCategory = "Совершенствование учебно-материальной базы";

    private readonly ExcelImportService _importService = new();
    private readonly SummaryCalculator _summaryCalculator = new();
    private readonly PlanRepository _repo;

    // --- Отчёт ---
    private readonly ReportService _reportService;
    private string? _lastDocxPath;
    private string? _lastPdfPath;

    // --- UI (левая панель) ---
    [ObservableProperty] private bool isLeftShown = true;

    public string ToggleLeftTitle => IsLeftShown ? "Свернуть меню" : "Показать меню";

    public void ToggleLeft()
    {
        IsLeftShown = !IsLeftShown;
        OnPropertyChanged(nameof(ToggleLeftTitle));
    }

    private static string ReportsDir =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "PlanMasterReports");

    private static string TemplatePath =>
        Path.Combine(AppContext.BaseDirectory, "Assets", "Templates", "ReportTemplate.docx");

    public MainWindowViewModel()
    {
        _repo = new PlanRepository(DbPath.GetDefaultPath());
        StatusText = "Готово";

        _reportService = new ReportService(TemplatePath);

        ResetMethodicalRows();

        _ = RefreshPlansAsync();
    }

    // Текущие таблицы (редактор)
    public ObservableCollection<PlanTable> Tables { get; } = new();
    public ObservableCollection<PlanRow> CurrentRows { get; } = new();
    public ObservableCollection<SummaryRow> SummaryRows { get; } = new();
    public ObservableCollection<MethodWorkRow> MethodicalProcessRows { get; } = new();
    public ObservableCollection<MethodWorkRow> MethodicalPublishingRows { get; } = new();
    public ObservableCollection<MethodWorkRow> MethodicalBaseRows { get; } = new();

    [ObservableProperty] private MethodWorkRow? selectedMethodicalProcessRow;
    [ObservableProperty] private MethodWorkRow? selectedMethodicalPublishingRow;
    [ObservableProperty] private MethodWorkRow? selectedMethodicalBaseRow;

    [ObservableProperty] private PlanTable? selectedTable;
    [ObservableProperty] private string statusText = "";

    public bool CanDeleteMethodicalProcessRow => SelectedMethodicalProcessRow != null;
    public bool CanDeleteMethodicalPublishingRow => SelectedMethodicalPublishingRow != null;
    public bool CanDeleteMethodicalBaseRow => SelectedMethodicalBaseRow != null;

    public string SelectedTableTitle
        => SelectedTable is null
            ? "Таблица не выбрана"
            : $"{SelectedTable.SheetName} — {SelectedTable.SemesterTitle}";

    partial void OnSelectedTableChanged(PlanTable? value)
    {
        CurrentRows.Clear();

        if (value?.Rows != null)
        {
            foreach (var r in value.Rows.OrderBy(r => r.RowOrder))
                CurrentRows.Add(r);
        }

        OnPropertyChanged(nameof(SelectedTableTitle));
    }

    // -----------------------------
    //  A) Импорт / Черновик
    // -----------------------------

    public async Task ImportFromExcelAsync(string filePath)
    {
        StatusText = "Импорт Excel…";

        var imported = _importService.Import(filePath);

        Tables.Clear();
        foreach (var t in imported.OrderBy(t => t.SheetName))
            Tables.Add(t);

        SelectedTable = Tables.FirstOrDefault();

        var summary = _summaryCalculator.CreateOrUpdate(existing: null, importedTables: imported);

        SummaryRows.Clear();
        foreach (var r in summary.Rows.OrderBy(r => r.RowOrder))
            SummaryRows.Add(r);

        ResetMethodicalRows();
        await _repo.SaveAllAsync(imported, summary, BuildMethodicalFromUi());

        CurrentPlanId = null;
        PlanName = "";

        _lastDocxPath = null;
        _lastPdfPath = null;

        StatusText = $"Импорт завершён: {Tables.Count} листов";
    }

    public async Task LoadFromDbAsync()
    {
        StatusText = "Загрузка из БД…";

        var (tables, summary, methodical) = await _repo.LoadAllAsync();

        Tables.Clear();
        foreach (var t in tables.OrderBy(t => t.SheetName))
            Tables.Add(t);

        SelectedTable = Tables.FirstOrDefault();

        SummaryRows.Clear();
        if (summary != null)
        {
            foreach (var r in summary.Rows.OrderBy(r => r.RowOrder))
                SummaryRows.Add(r);
        }

        LoadMethodical(methodical);

        LoadMethodical(methodical);

        CurrentPlanId = null;
        PlanName = "";

        _lastDocxPath = null;
        _lastPdfPath = null;

        StatusText = "Загружено (черновик)";
    }

    public async Task SaveToDbAsync()
    {
        StatusText = "Сохранение…";

        var recalced = BuildRecalcedSummaryFromUi();

        await _repo.SaveAllAsync(Tables.ToList(), recalced, BuildMethodicalFromUi());

        RefreshSummaryRows(recalced);

        StatusText = "Сохранено (черновик)";
    }

    private SummaryTable BuildSummaryFromUi()
    {
        return new SummaryTable
        {
            Rows = SummaryRows
                .OrderBy(r => r.RowOrder)
                .Select(r => new SummaryRow
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
                    IsTotalRow = r.IsTotalRow
                })
                .ToList()
        };
    }

    // -----------------------------
    //  B) Планы (много наборов)
    // -----------------------------

    public ObservableCollection<Plan> Plans { get; } = new();

    [ObservableProperty] private Plan? selectedPlan;
    [ObservableProperty] private string planName = "";
    [ObservableProperty] private int? currentPlanId;

    public bool CanSaveCurrentPlan => CurrentPlanId.HasValue;

    public async Task RefreshPlansAsync()
    {
        var list = await _repo.ListPlansAsync();

        Plans.Clear();
        foreach (var p in list)
            Plans.Add(p);
    }

    public async Task OpenSelectedPlanAsync()
    {
        if (SelectedPlan == null) return;

        StatusText = "Открытие плана…";

        var (tables, summary, methodical) = await _repo.LoadPlanAsync(SelectedPlan.Id);

        Tables.Clear();
        foreach (var t in tables.OrderBy(t => t.SheetName))
            Tables.Add(t);

        SelectedTable = Tables.FirstOrDefault();

        SummaryRows.Clear();
        if (summary != null)
        {
            foreach (var r in summary.Rows.OrderBy(r => r.RowOrder))
                SummaryRows.Add(r);
        }

        LoadMethodical(methodical);

        LoadMethodical(methodical);

        CurrentPlanId = SelectedPlan.Id;
        PlanName = SelectedPlan.Name;

        _lastDocxPath = null;
        _lastPdfPath = null;

        OnPropertyChanged(nameof(CanSaveCurrentPlan));

        StatusText = $"Открыт план: {PlanName}";
    }

    public async Task SaveAsNewPlanAsync()
    {
        StatusText = "Сохранение как новый план…";

        var recalced = BuildRecalcedSummaryFromUi();

        var newId = await _repo.SaveAsNewPlanAsync(PlanName, Tables.ToList(), recalced, BuildMethodicalFromUi());

        CurrentPlanId = newId;
        OnPropertyChanged(nameof(CanSaveCurrentPlan));

        RefreshSummaryRows(recalced);

        await RefreshPlansAsync();

        StatusText = "Сохранено как новый план";
    }

    public async Task SaveCurrentPlanAsync()
    {
        if (!CurrentPlanId.HasValue) return;

        StatusText = "Сохранение изменений плана…";

        var recalced = BuildRecalcedSummaryFromUi();

        await _repo.SavePlanAsync(CurrentPlanId.Value, PlanName, Tables.ToList(), recalced, BuildMethodicalFromUi());

        RefreshSummaryRows(recalced);

        await RefreshPlansAsync();

        StatusText = "План сохранён";
    }

    public async Task DeleteSelectedPlanAsync()
    {
        if (SelectedPlan == null) return;

        StatusText = "Удаление плана…";

        var id = SelectedPlan.Id;
        await _repo.DeletePlanAsync(id);

        await RefreshPlansAsync();

        if (CurrentPlanId == id)
        {
            CurrentPlanId = null;
            PlanName = "";
            OnPropertyChanged(nameof(CanSaveCurrentPlan));
        }

        StatusText = "План удалён";
    }

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

    public void AddMethodicalProcessRow()
        => MethodicalProcessRows.Add(new MethodWorkRow { Category = MethodicalProcessCategory });

    public void AddMethodicalPublishingRow()
        => MethodicalPublishingRows.Add(new MethodWorkRow { Category = MethodicalPublishingCategory });

    public void AddMethodicalBaseRow()
        => MethodicalBaseRows.Add(new MethodWorkRow { Category = MethodicalBaseCategory });

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

    // -----------------------------
    //  C) Отчёт (DOCX/PDF: создать, открыть, печать)
    // -----------------------------

    private bool HasDataForReport()
        => Tables.Count > 0 && SummaryRows.Count > 0;

    public Task GenerateReportDocxAsync()
    {
        if (!HasDataForReport())
        {
            StatusText = "Нет данных для отчёта: импортируйте Excel или откройте план.";
            return Task.CompletedTask;
        }

        if (!File.Exists(TemplatePath))
        {
            StatusText = $"Шаблон не найден: {TemplatePath}";
            return Task.CompletedTask;
        }

        try
        {
            StatusText = "Генерация DOCX…";
            Directory.CreateDirectory(ReportsDir);

            var outDocx = Path.Combine(ReportsDir, $"Отчет_{DateTime.Now:yyyyMMdd_HHmm}.docx");

            var recalced = BuildRecalcedSummaryFromUi();
            RefreshSummaryRows(recalced);

            _lastDocxPath = _reportService.GenerateDocx(
                outDocx,
                Tables.ToList(),
                recalced.Rows.ToList());

            _lastPdfPath = null;

            StatusText = $"DOCX готов: {_lastDocxPath}";
        }
        catch (Exception ex)
        {
            StatusText = "Ошибка генерации DOCX: " + ex.Message;
        }

        return Task.CompletedTask;
    }

    public async Task GenerateReportPdfAsync()
    {
        if (!HasDataForReport())
        {
            StatusText = "Нет данных для отчёта: импортируйте Excel или откройте план.";
            return;
        }

        if (!File.Exists(TemplatePath))
        {
            StatusText = $"Шаблон не найден: {TemplatePath}";
            return;
        }

        try
        {
            StatusText = "Генерация PDF…";
            Directory.CreateDirectory(ReportsDir);

            if (string.IsNullOrWhiteSpace(_lastDocxPath) || !File.Exists(_lastDocxPath))
                await GenerateReportDocxAsync();

            if (string.IsNullOrWhiteSpace(_lastDocxPath) || !File.Exists(_lastDocxPath))
                return;

            _lastPdfPath = _reportService.ConvertToPdfWithLibreOffice(_lastDocxPath, ReportsDir);

            StatusText = $"PDF готов: {_lastPdfPath}";
        }
        catch (Exception ex)
        {
            StatusText = "Ошибка генерации PDF: " + ex.Message;
        }
    }

    public void OpenLastDocx()
    {
        if (!string.IsNullOrWhiteSpace(_lastDocxPath) && File.Exists(_lastDocxPath))
        {
            _reportService.OpenFile(_lastDocxPath);
            StatusText = "Открыт DOCX";
        }
        else
        {
            StatusText = "DOCX ещё не создан.";
        }
    }

    public void OpenLastPdf()
    {
        if (!string.IsNullOrWhiteSpace(_lastPdfPath) && File.Exists(_lastPdfPath))
        {
            _reportService.OpenFile(_lastPdfPath);
            StatusText = "Открыт PDF";
        }
        else
        {
            StatusText = "PDF ещё не создан (нажмите 'PDF: Сгенерировать').";
        }
    }

    public async Task PrintLastDocxAsync()
    {
        await GenerateReportPdfAsync();

        if (!string.IsNullOrWhiteSpace(_lastPdfPath) && File.Exists(_lastPdfPath))
        {
            try
            {
                _reportService.PrintPdf(_lastPdfPath);
                StatusText = "Отправлено на печать (PDF)";
            }
            catch (Exception ex)
            {
                StatusText = "Ошибка печати: " + ex.Message;
            }
        }
        else
        {
            StatusText = "Не удалось подготовить PDF для печати.";
        }
    }

    public void PrintLastPdf()
    {
        if (!string.IsNullOrWhiteSpace(_lastPdfPath) && File.Exists(_lastPdfPath))
        {
            try
            {
                var sent = _reportService.PrintPdf(_lastPdfPath);
                StatusText = sent
                    ? "Отправлено на печать (PDF)"
                    : "Открыл PDF — нажмите 'Печать' в просмотрщике (Windows).";
            }
            catch (Exception ex)
            {
                StatusText = "Ошибка печати: " + ex.Message;
            }
        }
        else
        {
            StatusText = "PDF ещё не создан.";
        }
    }

}
