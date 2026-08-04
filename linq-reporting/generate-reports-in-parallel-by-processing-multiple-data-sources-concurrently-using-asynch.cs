using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static async Task Main()
    {
        // Ensure the working directory exists.
        Directory.CreateDirectory("Output");

        // Create two template documents programmatically.
        CreateTemplate("template1.docx");
        CreateTemplate("template2.docx");

        // Prepare sample data for each report.
        var model1 = new ReportModel
        {
            Title = "First Report",
            Description = "This is the description of the first report."
        };

        var model2 = new ReportModel
        {
            Title = "Second Report",
            Description = "This is the description of the second report."
        };

        // Generate reports in parallel.
        Task task1 = GenerateReportAsync(
            "template1.docx",
            Path.Combine("Output", "Report1.docx"),
            model1,
            "model");

        Task task2 = GenerateReportAsync(
            "template2.docx",
            Path.Combine("Output", "Report2.docx"),
            model2,
            "model");

        await Task.WhenAll(task1, task2);
    }

    // Creates a simple template with LINQ Reporting tags.
    private static void CreateTemplate(string fileName)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Report");
        builder.Writeln("Title: <<[model.Title]>>");
        builder.Writeln("Description: <<[model.Description]>>");

        doc.Save(fileName);
    }

    // Asynchronously loads a template, builds the report, and saves the result.
    private static async Task GenerateReportAsync(string templatePath, string outputPath, object model, string rootName)
    {
        await Task.Run(() =>
        {
            // Load the template document.
            var doc = new Document(templatePath);

            // Configure the reporting engine.
            var engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None;

            // Build the report using the provided model and root name.
            engine.BuildReport(doc, model, rootName);

            // Save the generated report.
            doc.Save(outputPath);
        });
    }
}

// Simple data model used by both templates.
public class ReportModel
{
    public string Title { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
}
