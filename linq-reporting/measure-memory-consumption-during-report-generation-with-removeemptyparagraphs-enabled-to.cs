using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some Aspose.Words features).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data model.
        ReportModel model = new()
        {
            Items = new()
        };
        for (int i = 1; i <= 1000; i++)
        {
            model.Items.Add(new Item
            {
                Index = i,
                Name = $"Item #{i}"
            });
        }

        // Create a template document programmatically.
        string templatePath = "template.docx";
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);

        builder.Writeln("=== Report Start ===");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Item <<[item.Index]>>: <<[item.Name]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("=== Report End ===");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        Document reportDoc = new(templatePath);

        // Measure memory before report generation.
        GC.Collect();
        GC.WaitForPendingFinalizers();
        long memoryBefore = GC.GetTotalMemory(true);

        // Configure and run the reporting engine with RemoveEmptyParagraphs enabled.
        ReportingEngine engine = new();
        engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;
        bool success = engine.BuildReport(reportDoc, model, "model");

        // Measure memory after report generation.
        GC.Collect();
        GC.WaitForPendingFinalizers();
        long memoryAfter = GC.GetTotalMemory(true);

        // Output results.
        Console.WriteLine($"Report generation success: {success}");
        Console.WriteLine($"Memory before: {memoryBefore:N0} bytes");
        Console.WriteLine($"Memory after : {memoryAfter:N0} bytes");
        Console.WriteLine($"Memory delta : {memoryAfter - memoryBefore:N0} bytes");

        // Save the generated report.
        string outputPath = Path.Combine("output", "ReportWithRemoveEmptyParagraphs.docx");
        Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);
        reportDoc.Save(outputPath);
        Console.WriteLine($"Report saved to: {outputPath}");
    }
}

// Data model classes.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = string.Empty;
}
