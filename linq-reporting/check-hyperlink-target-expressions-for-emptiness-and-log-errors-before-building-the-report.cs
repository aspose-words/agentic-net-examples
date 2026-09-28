using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Paths for the template and the generated report.
        string templatePath = Path.Combine(outputDir, "template.docx");
        string reportPath = Path.Combine(outputDir, "report.docx");

        // -----------------------------------------------------------------
        // 1. Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Add a title.
        builder.Writeln("Hyperlink Report");
        builder.Writeln();

        // Begin a foreach loop over Items.
        builder.Writeln("<<foreach [item in Items]>>");
        // Insert a link tag. If Url is empty, the link will be invalid – we will catch that before building.
        builder.Writeln("<<link [item.Url] [item.Text]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Prepare sample data, including an item with an empty hyperlink target.
        // -----------------------------------------------------------------
        ReportModel model = new ReportModel
        {
            Items = new List<ReportItem>
            {
                new ReportItem { Url = "https://example.com", Text = "Example Site" },
                new ReportItem { Url = "", Text = "Missing URL" }, // This should trigger an error log.
                new ReportItem { Url = "https://dotnet.microsoft.com", Text = ".NET Home" }
            }
        };

        // -----------------------------------------------------------------
        // 3. Validate hyperlink targets before building the report.
        // -----------------------------------------------------------------
        bool hasErrors = false;
        foreach (var item in model.Items)
        {
            if (string.IsNullOrWhiteSpace(item.Url))
            {
                Console.WriteLine($"Error: Hyperlink target is empty for display text \"{item.Text}\".");
                hasErrors = true;
            }
        }

        if (hasErrors)
        {
            Console.WriteLine("Report generation aborted due to validation errors.");
            return;
        }

        // -----------------------------------------------------------------
        // 4. Load the template and build the report.
        // -----------------------------------------------------------------
        Document loadedTemplate = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();

        // Build the report; root object name is "model" to match the template tags.
        engine.BuildReport(loadedTemplate, model, "model");

        // Save the generated report.
        loadedTemplate.Save(reportPath);
        Console.WriteLine($"Report generated successfully: {reportPath}");
    }
}

// ---------------------------------------------------------------------
// Data model definitions.
// ---------------------------------------------------------------------
public class ReportModel
{
    public List<ReportItem> Items { get; set; } = new();
}

public class ReportItem
{
    public string Url { get; set; } = string.Empty;
    public string Text { get; set; } = string.Empty;
}
