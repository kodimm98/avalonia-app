namespace PlanMaster.Models;

public class SummaryTable
{
    public int Id { get; set; }
    public int PlanId { get; set; }
    public DateTime CreatedAtUtc { get; set; } = DateTime.UtcNow;

    public List<SummaryRow> Rows { get; set; } = new();
}