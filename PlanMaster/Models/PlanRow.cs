namespace PlanMaster.Models;

public class PlanRow
{
    public int Id { get; set; }

    public int PlanTableId { get; set; }
    public PlanTable? PlanTable { get; set; }

    public int RowOrder { get; set; }

    public int? Number { get; set; }             // № п.п
    public string DisciplineName { get; set; } = "";
    public string FacultyGroup { get; set; } = "";

    // "Количество"
    public int? Course { get; set; }
    public int? Streams { get; set; }
    public int? Groups { get; set; }
    public int? Students { get; set; }

    // Часы по видам (если какие-то не нужны — можно потом убрать)
    public int? Lek { get; set; }
    public int? Pr { get; set; }
    public int? Lab { get; set; }
    public int? Ksr { get; set; }
    public int? Kp { get; set; }
    public int? Kr { get; set; }
    public int? KontrolRab { get; set; }
    public int? Zach { get; set; }
    public int? DifZach { get; set; }
    public int? Exz { get; set; }
    public int? GosExz { get; set; }
    public int? Gek { get; set; }
    public int? RukVkr { get; set; }
    public int? Rec { get; set; }
    public int? UchPr { get; set; }
    public int? PrPr { get; set; }
    public int? PredPr { get; set; }

    public int? Total { get; set; }              // "ВСЕГО" (колонка Y в твоём файле)
    public string? Note { get; set; }            // "Примечание"

    public bool IsSummary { get; set; }  // строка "Итого/Поручено/Выполнено"

}