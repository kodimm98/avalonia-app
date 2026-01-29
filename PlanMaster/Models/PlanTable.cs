namespace PlanMaster.Models;

public class PlanTable
{
    public int Id { get; set; }
    public int PlanId { get; set; } 

    public string SheetName { get; set; } = "";
    public string? SemesterTitle { get; set; } // "1 семестр" / "2 семестр"
    public DateTime ImportedAtUtc { get; set; } = DateTime.UtcNow;

    // Итоги (из строк "Итого за ... Поручено/Выполнено", колонка "ВСЕГО")
    public int? Sem1Plan { get; set; }
    public int? Sem1Fact { get; set; }
    public int? Sem2Plan { get; set; }
    public int? Sem2Fact { get; set; }
    public int? YearPlan { get; set; }
    public int? YearFact { get; set; }

    public List<PlanRow> Rows { get; set; } = new();
}