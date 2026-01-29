using System;
using System.Collections.Generic;

namespace PlanMaster.Models;

public class MethodWorkTable
{
    public int Id { get; set; }
    public int PlanId { get; set; }
    public DateTime CreatedAtUtc { get; set; } = DateTime.UtcNow;

    public List<MethodWorkRow> Rows { get; set; } = new();
}
