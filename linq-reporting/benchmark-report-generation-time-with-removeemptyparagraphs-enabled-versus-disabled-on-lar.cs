using System;
using System.Collections.Generic;
using System.Diagnostics;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportItem
{
    // Some items will have a null value to produce empty paragraphs after merging.
    public string? Value { get; set; }
}

public class ReportModel
{
    public List<ReportItem> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare a large data set (e.g., 10,000 items, half with null values).
        var model = new ReportModel();
        const int itemCount = 10000;
        for (int i = 0; i < itemCount; i++)
        {
            model.Items.Add(new ReportItem
            {
                Value = i % 2 == 0 ? $"Item {i}" : null   // Even indices have text, odd are null.
            });
        }

        // Create the LINQ Reporting template programmatically.
        const string templatePath = "Template.docx";
        CreateTemplate(templatePath);

        // Benchmark with RemoveEmptyParagraphs disabled.
        var timeWithoutRemoval = BenchmarkReport(templatePath, model, ReportBuildOptions.None, "Report_WithoutRemoveEmptyParagraphs.docx");

        // Benchmark with RemoveEmptyParagraphs enabled.
        var timeWithRemoval = BenchmarkReport(templatePath, model, ReportBuildOptions.RemoveEmptyParagraphs, "Report_WithRemoveEmptyParagraphs.docx");

        // Output the results.
        Console.WriteLine($"Report generation without RemoveEmptyParagraphs: {timeWithoutRemoval.TotalMilliseconds} ms");
        Console.WriteLine($"Report generation with    RemoveEmptyParagraphs: {timeWithRemoval.TotalMilliseconds} ms");
    }

    private static void CreateTemplate(string filePath)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Begin a foreach block over Items.
        builder.Writeln("<<foreach [item in Items]>>");
        // Each iteration writes the value; if the value is null the paragraph becomes empty.
        builder.Writeln("<<[item.Value]>>");
        // End the foreach block.
        builder.Writeln("<</foreach>>");

        doc.Save(filePath);
    }

    private static TimeSpan BenchmarkReport(string templatePath, ReportModel model, ReportBuildOptions options, string outputPath)
    {
        // Load the template.
        var doc = new Document(templatePath);

        // Configure the reporting engine.
        var engine = new ReportingEngine
        {
            Options = options
        };

        // Measure the time taken to build the report.
        var stopwatch = Stopwatch.StartNew();
        engine.BuildReport(doc, model, "model");
        stopwatch.Stop();

        // Save the generated document.
        doc.Save(outputPath);

        return stopwatch.Elapsed;
    }
}
