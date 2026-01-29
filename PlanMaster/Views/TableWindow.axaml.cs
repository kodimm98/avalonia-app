using Avalonia.Controls;
using Avalonia.Interactivity;

namespace PlanMaster.Views;

public partial class TableWindow : Window
{
    public TableWindow()
    {
        InitializeComponent();
    }

    private void Close_Click(object? sender, RoutedEventArgs e) => Close();
}