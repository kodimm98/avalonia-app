using System.IO;
using System;

namespace PlanMaster.Data;

public static class DbPath
{
    public static string GetDefaultPath()
    {
        var dir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "PlanMaster");

        Directory.CreateDirectory(dir);
        return Path.Combine(dir, "planmaster.db");
    }
}