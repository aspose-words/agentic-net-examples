using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Sample data properties. Optional may be null, causing an empty paragraph after rendering.
    public string Name { get; set; } = "John Doe";
    public string? Optional { get; set; } = null;
}

public class Program
{
    public static void Main()
    {
        // Prepare a simple template document with a placeholder that can become empty.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Paragraph with a regular value.
        builder.Writeln("Name: <<[model.Name]>>");
        // Paragraph that contains only a tag which may resolve to an empty string.
        builder.Writeln("<<[model.Optional]>>");
        // Final paragraph to show the document after processing.
        builder.Writeln("End of report.");

        // Save the template locally (optional, shown for completeness).
        const string templatePath = "template.docx";
        template.Save(templatePath);

        // Load the template back (demonstrates the load step).
        Document doc = new Document(templatePath);

        // Create the data source.
        ReportModel model = new ReportModel();

        // Configure the reporting engine to remove empty paragraphs after tags are processed.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.RemoveEmptyParagraphs
        };

        // Build the report. The root object name must match the tag prefix used in the template.
        engine.BuildReport(doc, model, "model");

        // Save the final document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);

        // Indicate completion (no interactive prompts).
        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}
