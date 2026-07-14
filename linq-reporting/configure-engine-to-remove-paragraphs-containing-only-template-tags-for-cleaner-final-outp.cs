using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Non‑nullable property – initialized to avoid warnings.
    public string Name { get; set; } = string.Empty;

    // Nullable property – may be null, causing the tag to become empty.
    public string? EmptyTag { get; set; }
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Name = "John Doe",
            EmptyTag = null // This will result in an empty tag after processing.
        };

        // Create a template document programmatically.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Paragraph that will always contain data.
        builder.Writeln("Customer: <<[model.Name]>>");

        // Paragraph that contains only a tag which resolves to an empty value.
        builder.Writeln("<<[model.EmptyTag]>>");

        // Configure the reporting engine to remove empty paragraphs.
        var engine = new ReportingEngine
        {
            Options = ReportBuildOptions.RemoveEmptyParagraphs
        };

        // Build the report using the model as the root data source.
        engine.BuildReport(doc, model, "model");

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Save the resulting document.
        string outputPath = Path.Combine(outputDir, "Report.docx");
        doc.Save(outputPath);

        // Inform the user where the file was saved.
        Console.WriteLine($"Report generated: {outputPath}");
    }
}
