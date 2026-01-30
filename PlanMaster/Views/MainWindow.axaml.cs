using System.Linq;
using Avalonia.Controls;
using Avalonia.Interactivity;
using PlanMaster.ViewModels;
using Avalonia.Platform.Storage;



namespace PlanMaster.Views;

public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
        DataContext ??= new MainWindowViewModel();
    }

    private async void ImportExcel_Click(object? sender, RoutedEventArgs e)
    {
        if (DataContext is not MainWindowViewModel vm) return;

        var files = await StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
        {
            Title = "Выберите Excel файл",
            AllowMultiple = false,
            FileTypeFilter = new[]
            {
                new FilePickerFileType("Excel") { Patterns = new[] { "*.xlsx" } }
            }
        });

        var file = files.FirstOrDefault();
        if (file == null) return;

        await vm.ImportFromExcelAsync(file.Path.LocalPath);
    }

    private async void LoadDb_Click(object? sender, RoutedEventArgs e)
    {
        if (DataContext is MainWindowViewModel vm)
            await vm.LoadFromDbAsync();
    }

    private async void SaveDb_Click(object? sender, RoutedEventArgs e)
    {
        if (DataContext is MainWindowViewModel vm)
            await vm.SaveToDbAsync();
    }

    private async void RefreshPlans_Click(object? sender, RoutedEventArgs e)
{
    if (DataContext is MainWindowViewModel vm)
        await vm.RefreshPlansAsync();
}

private async void OpenPlan_Click(object? sender, RoutedEventArgs e)
{
    if (DataContext is MainWindowViewModel vm)
        await vm.OpenSelectedPlanAsync();
}

private async void SaveAsNewPlan_Click(object? sender, RoutedEventArgs e)
{
    if (DataContext is MainWindowViewModel vm)
        await vm.SaveAsNewPlanAsync();
}

private async void SaveCurrentPlan_Click(object? sender, RoutedEventArgs e)
{
    if (DataContext is MainWindowViewModel vm)
        await vm.SaveCurrentPlanAsync();
}

private async void DeletePlan_Click(object? sender, RoutedEventArgs e)
{
    if (DataContext is MainWindowViewModel vm)
        await vm.DeleteSelectedPlanAsync();
}


private async void ReportDocx_Click(object? sender, RoutedEventArgs e)
{
    await ((MainWindowViewModel)DataContext!).GenerateReportDocxAsync();
}

private async void ReportPdf_Click(object? sender, RoutedEventArgs e)
{
    await ((MainWindowViewModel)DataContext!).GenerateReportPdfAsync();
}

private void OpenLastDocx_Click(object? sender, RoutedEventArgs e)
{
    ((MainWindowViewModel)DataContext!).OpenLastDocx();
}

private async void PrintLastDocx_Click(object? sender, RoutedEventArgs e)
{
    await ((MainWindowViewModel)DataContext!).PrintLastDocxAsync();
}

private void OpenLastPdf_Click(object? sender, RoutedEventArgs e)
{
    ((MainWindowViewModel)DataContext!).OpenLastPdf();
}

private void PrintLastPdf_Click(object? sender, RoutedEventArgs e)
{
    ((MainWindowViewModel)DataContext!).PrintLastPdf();
}

private void ToggleLeft_Click(object? sender, Avalonia.Interactivity.RoutedEventArgs e) { if (DataContext is PlanMaster.ViewModels.MainWindowViewModel vm) vm.ToggleLeft(); }

        private void OpenInWindow_Click(object? sender, RoutedEventArgs e)
    {
        if (DataContext is not MainWindowViewModel vm) return;
        if (vm.SelectedTable is null) return;

        var w = new TableWindow
        {
            DataContext = new TableWindowViewModel(vm.SelectedTable),
            WindowState = WindowState.Maximized
        };

        w.Show();
    }

    private void PlansMenu_Click(object? sender, RoutedEventArgs e)
    {
        if (DataContext is not MainWindowViewModel vm) return;

        var window = new PlansWindow
        {
            DataContext = vm
        };

        window.Show(this);
    }

    private void AddMethodicalProcessRow_Click(object? sender, RoutedEventArgs e)
    {
        if (DataContext is MainWindowViewModel vm)
            vm.AddMethodicalProcessRow();
    }

    private void AddMethodicalPublishingRow_Click(object? sender, RoutedEventArgs e)
    {
        if (DataContext is MainWindowViewModel vm)
            vm.AddMethodicalPublishingRow();
    }

    private void AddMethodicalBaseRow_Click(object? sender, RoutedEventArgs e)
    {
        if (DataContext is MainWindowViewModel vm)
            vm.AddMethodicalBaseRow();
    }
}
