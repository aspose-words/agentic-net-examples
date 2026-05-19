using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public int Index { get; set; } = 0;
    public string Name { get; set; } = string.Empty;
}

public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Program
{
    // Asynchronous method that builds a report for a given model.
    private static async Task GenerateReportAsync(int reportId, ReportModel model, string templatePath)
    {
        // Load the previously saved template.
        var doc = new Document(templatePath);

        // Create and configure the reporting engine.
        var engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };

        // Build the report using the model as the root data source named "model".
        engine.BuildReport(doc, model, "model");

        // Ensure the output directory exists.
        var outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Save the generated report.
        var outputPath = Path.Combine(outputDir, $"Report_{reportId}.docx");
        doc.Save(outputPath);
    }

    // Entry point of the console application.
    public static async Task Main()
    {
        // Define paths for the template.
        var templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");

        // -----------------------------------------------------------------
        // Step 1: Create the template document programmatically.
        // -----------------------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Simple heading.
        builder.Writeln("Report of Items:");

        // LINQ Reporting foreach tag to iterate over Items collection.
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("- <<[item.Index]>>: <<[item.Name]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Step 2: Prepare multiple data models for parallel processing.
        // -----------------------------------------------------------------
        var models = new List<ReportModel>();
        for (int i = 1; i <= 3; i++)
        {
            var model = new ReportModel
            {
                Items = new List<Item>
                {
                    new() { Index = 1, Name = $"Alpha_{i}" },
                    new() { Index = 2, Name = $"Beta_{i}" },
                    new() { Index = 3, Name = $"Gamma_{i}" }
                }
            };
            models.Add(model);
        }

        // -----------------------------------------------------------------
        // Step 3: Generate reports concurrently using asynchronous tasks.
        // -----------------------------------------------------------------
        var tasks = new List<Task>();
        for (int i = 0; i < models.Count; i++)
        {
            int reportId = i + 1; // Capture loop variable.
            var model = models[i];
            tasks.Add(Task.Run(() => GenerateReportAsync(reportId, model, templatePath)));
        }

        // Await all report generation tasks.
        await Task.WhenAll(tasks);
    }
}
