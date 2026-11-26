using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare folders
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Create template document
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert LINQ Reporting tags
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Item: <<[item.Name]>> - Value: <<[item.Value]>>");
        builder.Writeln("<</foreach>>");

        // Save template to disk
        string templatePath = Path.Combine(outputDir, "template.docx");
        template.Save(templatePath);

        // Load template for reporting
        Document doc = new Document(templatePath);
        DocumentBuilder reportBuilder = new DocumentBuilder(doc); // Not strictly needed but keeps builder available if further edits required

        // Prepare sample data
        ReportModel model = new ReportModel
        {
            Items = new List<Item>()
        };
        for (int i = 1; i <= 1000; i++)
        {
            model.Items.Add(new Item { Name = $"Item{i}", Value = i * 10 });
        }

        // Measure memory before report generation
        long memoryBefore = GC.GetTotalMemory(true);

        // Build report with RemoveEmptyParagraphs option
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;
        engine.BuildReport(doc, model, "model");

        // Measure memory after report generation
        long memoryAfter = GC.GetTotalMemory(true);
        long memoryDiff = memoryAfter - memoryBefore;

        // Save the generated report
        string outputPath = Path.Combine(outputDir, "report.docx");
        doc.Save(outputPath);

        // Output memory consumption information
        Console.WriteLine($"Memory before: {memoryBefore:N0} bytes");
        Console.WriteLine($"Memory after : {memoryAfter:N0} bytes");
        Console.WriteLine($"Memory diff  : {memoryDiff:N0} bytes");
        Console.WriteLine($"Report saved to: {outputPath}");
    }
}

// Data model classes
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = string.Empty;
    public int Value { get; set; }
}
