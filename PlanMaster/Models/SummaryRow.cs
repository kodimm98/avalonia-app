namespace PlanMaster.Models;

public class SummaryRow
{
    public int Id { get; set; }

    public int SummaryTableId { get; set; }
    public SummaryTable? SummaryTable { get; set; }

    public int RowOrder { get; set; }

    public string Code { get; set; } = "";      // "1", "2", "2.1", ...
    public string WorkName { get; set; } = "";  // "Учебная работа" и т.п.

    public int Sem1Plan { get; set; }
    public int Sem1Fact { get; set; }
    public int Sem2Plan { get; set; }
    public int Sem2Fact { get; set; }
    public int YearPlan { get; set; }
    public int YearFact { get; set; }

    public bool IsTotalRow { get; set; }        // строка "Итого:"
}