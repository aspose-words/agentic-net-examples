using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Initialize the collection to avoid nullable warnings.
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = string.Empty;
}

public class Program
{
    public static void Main()
    {
        // Enable reflection optimization for maximum performance.
        ReportingEngine.UseReflectionOptimization = true;

        // Prepare sample data with a large number of items.
        var model = new ReportModel();
        for (int i = 1; i <= 1000; i++)
        {
            model.Items.Add(new Item { Index = i, Name = $"Item {i}" });
        }

        // Create a template document programmatically.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Insert LINQ Reporting tags.
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("<<[item.Index]>> - <<[item.Name]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(templatePath);

        // Load the template and build the report.
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        reportDoc.Save(outputPath);
    }
}
