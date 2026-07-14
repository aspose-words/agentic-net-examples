using System;
using System.Collections.Generic;
using System.Diagnostics;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Paths for the template and generated reports.
        const string templatePath = "template.docx";
        const string reportWithOptionPath = "Report_With_RemoveEmptyParagraphs.docx";
        const string reportWithoutOptionPath = "Report_Without_RemoveEmptyParagraphs.docx";

        // -----------------------------------------------------------------
        // 1. Create a LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Simple header.
        builder.Writeln("=== LINQ Reporting Benchmark ===");
        // Begin a foreach loop over Items.
        builder.Writeln("<<foreach [item in Items]>>");
        // Each iteration writes two fields; the second field may be empty for some items.
        builder.Writeln("Item <<[item.Index]>>: <<[item.Text]>>");
        // End the loop.
        builder.Writeln("<</foreach>>");
        // Footer.
        builder.Writeln("=== End of Report ===");

        // Save the template to disk so it can be re‑loaded for each benchmark run.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Prepare a large data source.
        // -----------------------------------------------------------------
        var model = new ReportModel
        {
            Items = GenerateItems(5000) // Adjust the count for desired size.
        };

        // -----------------------------------------------------------------
        // 3. Benchmark with RemoveEmptyParagraphs enabled.
        // -----------------------------------------------------------------
        var timeWithOption = BenchmarkReportGeneration(
            templatePath,
            model,
            ReportBuildOptions.RemoveEmptyParagraphs,
            reportWithOptionPath);

        // -----------------------------------------------------------------
        // 4. Benchmark with RemoveEmptyParagraphs disabled.
        // -----------------------------------------------------------------
        var timeWithoutOption = BenchmarkReportGeneration(
            templatePath,
            model,
            ReportBuildOptions.None,
            reportWithoutOptionPath);

        // -----------------------------------------------------------------
        // 5. Output the results.
        // -----------------------------------------------------------------
        Console.WriteLine($"Report generation with RemoveEmptyParagraphs: {timeWithOption.TotalMilliseconds} ms");
        Console.WriteLine($"Report generation without RemoveEmptyParagraphs: {timeWithoutOption.TotalMilliseconds} ms");
    }

    // Generates a list of items; every 10th item has an empty Text to trigger empty paragraph removal.
    private static List<Item> GenerateItems(int count)
    {
        var items = new List<Item>(count);
        for (int i = 1; i <= count; i++)
        {
            items.Add(new Item
            {
                Index = i,
                Text = i % 10 == 0 ? string.Empty : $"Sample text for item {i}"
            });
        }
        return items;
    }

    // Runs the report generation once and returns the elapsed time.
    private static TimeSpan BenchmarkReportGeneration(
        string templatePath,
        ReportModel model,
        ReportBuildOptions options,
        string outputPath)
    {
        // Load a fresh copy of the template for each run.
        Document doc = new Document(templatePath);

        // Configure the reporting engine.
        ReportingEngine engine = new ReportingEngine
        {
            Options = options
        };

        // Measure the build time.
        Stopwatch sw = Stopwatch.StartNew();
        bool success = engine.BuildReport(doc, model, "model");
        sw.Stop();

        // Save the generated report (optional, but ensures the document is fully built).
        doc.Save(outputPath);

        // The success flag is only meaningful when InlineErrorMessages is set; we ignore it here.
        return sw.Elapsed;
    }
}

// ---------------------------------------------------------------------
// Data model classes – must be public with public properties.
// ---------------------------------------------------------------------
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public int Index { get; set; }
    public string Text { get; set; } = string.Empty;
}
