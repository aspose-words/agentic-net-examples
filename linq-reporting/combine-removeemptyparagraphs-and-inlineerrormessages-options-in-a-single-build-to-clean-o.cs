using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare a simple data model.
        var model = new ReportModel
        {
            Name = "John Doe",
            Empty = string.Empty // This will produce an empty paragraph.
            // Note: No property named NonExistent – accessing it will cause an error.
        };

        // Create a template document programmatically.
        var templatePath = "template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Normal field.
        builder.Writeln("Customer: <<[model.Name]>>");
        // This tag references a missing member and will generate an inline error message.
        builder.Writeln("Missing: <<[model.NonExistent]>>");
        // This tag resolves to an empty string; the paragraph should be removed.
        builder.Writeln("Empty: <<[model.Empty]>>");
        // A paragraph that contains only an empty tag; it should be removed entirely.
        builder.Writeln("<<[model.Empty]>>");

        // Save the template to disk.
        doc.Save(templatePath);

        // Load the template back for reporting.
        var loadedDoc = new Document(templatePath);

        // Configure the reporting engine with both options.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.RemoveEmptyParagraphs | ReportBuildOptions.InlineErrorMessages;

        // Build the report. The returned flag indicates success when InlineErrorMessages is set.
        bool success = engine.BuildReport(loadedDoc, model, "model");

        // Save the resulting document.
        var outputPath = "output.docx";
        loadedDoc.Save(outputPath);

        // Output simple status to the console (no user interaction required).
        Console.WriteLine($"Report build success: {success}");
        Console.WriteLine($"Output saved to: {outputPath}");
    }
}

// Data model used by the template.
public class ReportModel
{
    // Initialized to avoid nullable warnings.
    public string Name { get; set; } = string.Empty;
    public string Empty { get; set; } = string.Empty;
}
