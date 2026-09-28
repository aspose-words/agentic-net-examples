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
        // Register code page provider for CSV handling.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create sample CSV data.
        string csvPath = "data.csv";
        File.WriteAllText(csvPath,
@"Category,Item
Fruits,Apple
Fruits,Banana
Fruits,Orange
Vegetables,Carrot
Vegetables,Tomato
Vegetables,Spinach
Grains,Rice
Grains,Wheat");

        // Load CSV and group by Category.
        var lines = File.ReadAllLines(csvPath);
        var data = lines.Skip(1)
                        .Select(l => l.Split(','))
                        .Select(parts => new { Category = parts[0].Trim(), Item = parts[1].Trim() })
                        .ToList();

        var model = new ReportModel
        {
            Categories = data.GroupBy(d => d.Category)
                            .Select(g => new Category
                            {
                                Name = g.Key,
                                Items = g.Select(x => x.Item).ToList()
                            })
                            .ToList()
        };

        // Create the template document programmatically.
        string templatePath = "template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Grouped Items Report");
        builder.Writeln();

        // Outer foreach for categories.
        builder.Writeln("<<foreach [cat in Categories]>>");
        builder.ListFormat.ApplyBulletDefault();
        builder.Writeln("<<[cat.Name]>>");

        // Inner foreach for items (indented bullet).
        builder.ListFormat.ListLevelNumber = 1; // Indent one level.
        builder.Writeln("<<foreach [item in cat.Items]>>");
        builder.Writeln("<<[item]>>");
        builder.Writeln("<</foreach>>");

        // Reset indentation and close outer foreach.
        builder.ListFormat.ListLevelNumber = 0;
        builder.Writeln("<</foreach>>");
        builder.ListFormat.RemoveNumbers();

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        var doc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;

        // Build the report.
        engine.BuildReport(doc, model, "model");

        // Save the final document.
        string outputPath = Path.Combine("output", "report.docx");
        Directory.CreateDirectory("output");
        doc.Save(outputPath);
    }
}

// Data model classes.
public class ReportModel
{
    public List<Category> Categories { get; set; } = new();
}

public class Category
{
    public string Name { get; set; } = "";
    public List<string> Items { get; set; } = new();
}
