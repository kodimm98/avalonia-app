namespace PlanMaster.Models;

public class MethodWorkRow
{
    public int Id { get; set; }

    public int MethodWorkTableId { get; set; }
    public MethodWorkTable? MethodWorkTable { get; set; }

    public int RowOrder { get; set; }
    public string Category { get; set; } = "";
    public string WorkName { get; set; } = "";
    public int? TimeHours { get; set; }
    public string Deadline { get; set; } = "";
    public string CompletionNote { get; set; } = "";
}
