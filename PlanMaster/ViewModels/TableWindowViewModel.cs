using System.Collections.ObjectModel;
using System.Linq;
using CommunityToolkit.Mvvm.ComponentModel;
using PlanMaster.Models;

namespace PlanMaster.ViewModels;

public partial class TableWindowViewModel : ObservableObject
{
    public string Title { get; }
    public ObservableCollection<PlanRow> Rows { get; } = new();

    public TableWindowViewModel(PlanTable table)
    {
        Title = $"{table.SheetName} — {table.SemesterTitle}";

        foreach (var r in table.Rows.OrderBy(r => r.RowOrder))
            Rows.Add(r);
    }
}