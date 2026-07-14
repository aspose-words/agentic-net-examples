using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // File paths (relative to the working directory).
        string csvPath = "data.csv";
        string templatePath = "template.docx";
        string outputPath = "report.docx";

        // -----------------------------------------------------------------
        // 1. Create sample CSV data.
        // -----------------------------------------------------------------
        // Header: Category,Item
        var csvLines = new[]
        {
            "Category,Item",
            "Fruits,Apple",
            "Fruits,Banana",
            "Fruits,Orange",
            "Vegetables,Carrot",
            "Vegetables,Potato",
            "Vegetables,Tomato"
        };
        File.WriteAllLines(csvPath, csvLines);

        // -----------------------------------------------------------------
        // 2. Load CSV and build a hierarchical model.
        // -----------------------------------------------------------------
        var groups = File.ReadAllLines(csvPath)
                         .Skip(1) // Skip header.
                         .Select(l => l.Split(','))
                         .Where(p => p.Length == 2)
                         .Select(p => new { Category = p[0].Trim(), Item = p[1].Trim() })
                         .GroupBy(x => x.Category)
                         .Select(g => new Group
                         {
                             Category = g.Key,
                             Items = g.Select(x => x.Item).ToList()
                         })
                         .ToList();

        var model = new ReportModel { Groups = groups };

        // -----------------------------------------------------------------
        // 3. Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Begin outer foreach over groups.
        builder.Writeln("<<foreach [group in Groups]>>");

        // Apply first‑level bullet (category).
        List list = templateDoc.Lists.Add(ListTemplate.BulletDefault);
        builder.ListFormat.List = list;
        builder.ListFormat.ListLevelNumber = 0;
        builder.Writeln("<<[group.Category]>>");

        // Begin inner foreach over items.
        builder.Writeln("<<foreach [item in group.Items]>>");

        // Apply second‑level bullet (item).
        builder.ListFormat.ListLevelNumber = 1;
        builder.Writeln("<<[item]>>");

        // End inner foreach.
        builder.Writeln("<</foreach>>");

        // End outer foreach.
        builder.Writeln("<</foreach>>");

        // Reset list formatting for any following content.
        builder.ListFormat.RemoveNumbers();

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 4. Build the report using the ReportingEngine.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // -----------------------------------------------------------------
        // 5. Save the generated report.
        // -----------------------------------------------------------------
        reportDoc.Save(outputPath);
    }
}

// ---------------------------------------------------------------------
// Data model classes (public, non‑nullable properties are initialized).
// ---------------------------------------------------------------------
public class ReportModel
{
    public List<Group> Groups { get; set; } = new();
}

public class Group
{
    public string Category { get; set; } = string.Empty;
    public List<string> Items { get; set; } = new();
}
