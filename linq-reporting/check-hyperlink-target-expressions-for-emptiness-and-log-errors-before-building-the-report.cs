using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Hyperlink target (URL or bookmark name). Initialized to empty string to avoid nullable warnings.
    public string Url { get; set; } = string.Empty;

    // Display text for the hyperlink.
    public string Text { get; set; } = string.Empty;
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data with one valid and one empty hyperlink target.
        var models = new List<ReportModel>
        {
            new ReportModel { Url = "https://www.example.com", Text = "Example Site" },
            new ReportModel { Url = "", Text = "Missing URL" } // This entry should trigger an error log.
        };

        // Create a template document programmatically.
        const string templatePath = "Template.docx";
        CreateTemplate(templatePath);

        // Load the template for reporting.
        Document templateDoc = new Document(templatePath);

        // Iterate over the data items, validate hyperlink targets, and log errors.
        foreach (var model in models)
        {
            if (string.IsNullOrWhiteSpace(model.Url))
            {
                Console.WriteLine($"Error: Hyperlink target is empty for display text \"{model.Text}\".");
                // Continue processing other items; the report will still be generated.
            }

            // Build the report for the current model.
            ReportingEngine engine = new ReportingEngine();
            // Use InlineErrorMessages to capture any template parsing issues.
            engine.Options = ReportBuildOptions.InlineErrorMessages;

            // The root object name in the template is "model".
            bool success = engine.BuildReport(templateDoc, model, "model");

            // Save the generated report with a distinct filename.
            string outputPath = $"Report_{Guid.NewGuid():N}.docx";
            templateDoc.Save(outputPath);
            Console.WriteLine($"Report generated: {outputPath} (Success: {success})");
        }
    }

    private static void CreateTemplate(string filePath)
    {
        // The template contains a LINQ Reporting link tag.
        // <<link [model.Url] [model.Text]>>
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("<<link [model.Url] [model.Text]>>");
        doc.Save(filePath);
    }
}
