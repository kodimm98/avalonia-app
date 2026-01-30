using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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

    [ObservableProperty] private PlanTable? selectedTable;
    [ObservableProperty] private string statusText = "";

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

        CurrentPlanId = null;
        PlanName = "";

        _lastDocxPath = null;
        _lastPdfPath = null;

        StatusText = "Загружено (черновик)";
    }

    public async Task SaveToDbAsync()
    {
        StatusText = "Сохранение…";

        var summary = BuildSummaryFromUi();
        var recalced = _summaryCalculator.CreateOrUpdate(summary, Tables.ToList());

        await _repo.SaveAllAsync(Tables.ToList(), recalced, BuildMethodicalFromUi());

        SummaryRows.Clear();
        foreach (var r in recalced.Rows.OrderBy(r => r.RowOrder))
            SummaryRows.Add(r);

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

        var summary = BuildSummaryFromUi();
        var recalced = _summaryCalculator.CreateOrUpdate(summary, Tables.ToList());

        var newId = await _repo.SaveAsNewPlanAsync(PlanName, Tables.ToList(), recalced, BuildMethodicalFromUi());

        CurrentPlanId = newId;
        OnPropertyChanged(nameof(CanSaveCurrentPlan));

        await RefreshPlansAsync();

        StatusText = "Сохранено как новый план";
    }

    public async Task SaveCurrentPlanAsync()
    {
        if (!CurrentPlanId.HasValue) return;

        StatusText = "Сохранение изменений плана…";

        var summary = BuildSummaryFromUi();
        var recalced = _summaryCalculator.CreateOrUpdate(summary, Tables.ToList());

        await _repo.SavePlanAsync(CurrentPlanId.Value, PlanName, Tables.ToList(), recalced, BuildMethodicalFromUi());

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

            _lastDocxPath = _reportService.GenerateDocx(
                outDocx,
                Tables.ToList(),
                SummaryRows.ToList());

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
