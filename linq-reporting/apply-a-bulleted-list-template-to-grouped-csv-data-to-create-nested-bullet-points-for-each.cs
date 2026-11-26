using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for possible CSV encoding handling.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample CSV data.
        const string csvPath = "data.csv";
        File.WriteAllText(csvPath, @"Category,Item
Fruits,Apple
Fruits,Banana
Fruits,Orange
Vegetables,Carrot
Vegetables,Tomato
Vegetables,Potato");

        // Load CSV into objects.
        List<Record> records = new();
        foreach (var line in File.ReadAllLines(csvPath).Skip(1))
        {
            var parts = line.Split(',');
            if (parts.Length >= 2)
            {
                records.Add(new Record
                {
                    Category = parts[0].Trim(),
                    Item = parts[1].Trim()
                });
            }
        }

        // Group records by category.
        List<CategoryGroup> groups = records
            .GroupBy(r => r.Category)
            .Select(g => new CategoryGroup
            {
                Category = g.Key,
                Items = g.Select(r => r.Item).ToList()
            })
            .ToList();

        // Create the LINQ Reporting template.
        const string templatePath = "template.docx";
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);

        // Begin bullet list.
        builder.ListFormat.ApplyBulletDefault();

        // Outer foreach for categories.
        builder.Writeln("<<foreach [group in Groups]>>");
        // Category bullet.
        builder.Writeln("<<[group.Category]>>");

        // Indent for sub‑items.
        builder.ListFormat.ListIndent();

        // Inner foreach for items.
        builder.Writeln("<<foreach [item in group.Items]>>");
        builder.Writeln("<<[item]>>");
        builder.Writeln("<</foreach>>");

        // Outdent back to category level.
        builder.ListFormat.ListOutdent();

        // End outer foreach.
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new(templatePath);

        // Build the report.
        ReportingEngine engine = new();
        engine.BuildReport(reportDoc, new ReportModel { Groups = groups }, "model");

        // Save the final document.
        const string outputPath = "output.docx";
        reportDoc.Save(outputPath);
    }
}

// Data model for a CSV record.
public class Record
{
    public string Category { get; set; } = "";
    public string Item { get; set; } = "";
}

// Grouped data model used by the template.
public class CategoryGroup
{
    public string Category { get; set; } = "";
    public List<string> Items { get; set; } = new();
}

// Root model passed to the ReportingEngine.
public class ReportModel
{
    public List<CategoryGroup> Groups { get; set; } = new();
}
