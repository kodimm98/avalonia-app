using System;

namespace PlanMaster.Models;

public class Plan
{
    public int Id { get; set; }
    public string Name { get; set; } = "";
    public DateTime CreatedAtUtc { get; set; }
    public DateTime UpdatedAtUtc { get; set; }
}