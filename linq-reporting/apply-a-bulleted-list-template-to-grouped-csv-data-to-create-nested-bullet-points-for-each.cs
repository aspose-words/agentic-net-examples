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
        // Register code page provider for CSV encoding support.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample CSV data.
        string csvPath = "data.csv";
        File.WriteAllText(csvPath,
            "Category,Item\r\n" +
            "Fruits,Apple\r\n" +
            "Fruits,Banana\r\n" +
            "Vegetables,Carrot\r\n" +
            "Vegetables,Tomato\r\n" +
            "Fruits,Orange");

        // Load CSV and group by Category.
        var groups = File.ReadAllLines(csvPath)
                         .Skip(1) // Skip header
                         .Select(line => line.Split(','))
                         .Where(parts => parts.Length >= 2)
                         .Select(parts => new { Category = parts[0].Trim(), Item = parts[1].Trim() })
                         .GroupBy(x => x.Category)
                         .Select(g => new CategoryGroup
                         {
                             Category = g.Key,
                             Items = g.Select(x => x.Item).ToList()
                         })
                         .ToList();

        // Build the root model.
        var model = new ReportModel
        {
            Categories = groups
        };

        // Create the LINQ Reporting template programmatically.
        string templatePath = "template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Begin outer foreach over categories.
        builder.Writeln("<<foreach [cat in model.Categories]>>");

        // First-level bullet for category name.
        builder.ListFormat.ApplyBulletDefault();
        builder.Writeln("<<[cat.Category]>>");

        // Indent for second-level bullets (items).
        builder.ListFormat.ListIndent();

        // Begin inner foreach over items.
        builder.Writeln("<<foreach [itm in cat.Items]>>");
        builder.Writeln("<<[itm]>>");
        builder.Writeln("<</foreach>>");

        // Outdent back to first level.
        builder.ListFormat.ListOutdent();

        // End outer foreach.
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        var doc = new Document(templatePath);

        // Build the report using LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the final report.
        string outputPath = "report.docx";
        doc.Save(outputPath);
    }
}

// Root model class.
public class ReportModel
{
    public List<CategoryGroup> Categories { get; set; } = new();
}

// Category group class.
public class CategoryGroup
{
    public string Category { get; set; } = "";
    public List<string> Items { get; set; } = new();
}
