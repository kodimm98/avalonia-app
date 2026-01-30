using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.Interactivity;
using PlanMaster.ViewModels;

namespace PlanMaster.Views;

public partial class PlansWindow : Window
{
    public PlansWindow()
    {
        InitializeComponent();
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

    private void Close_Click(object? sender, RoutedEventArgs e)
    {
        Close();
    }
}
