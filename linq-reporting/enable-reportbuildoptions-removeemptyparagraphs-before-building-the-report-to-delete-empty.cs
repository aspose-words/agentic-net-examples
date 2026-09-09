using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a blank document and a builder to add content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph with a value that will be replaced.
        builder.Writeln("Hello <<[model.Name]>>");

        // This tag resolves to an empty string, resulting in an empty paragraph.
        builder.Writeln("<<[model.Empty]>>");

        // Add another paragraph after the empty one.
        builder.Writeln("World");

        // Prepare the data model.
        ReportModel model = new ReportModel
        {
            Name = "John",
            Empty = string.Empty // will produce an empty paragraph.
        };

        // Configure the reporting engine to remove empty paragraphs.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.RemoveEmptyParagraphs
        };

        // Build the report using the model as the root object named "model".
        engine.BuildReport(doc, model, "model");

        // Save the resulting document.
        doc.Save("ReportWithRemovedEmptyParagraphs.docx");
    }
}

// Simple data model used by the template.
public class ReportModel
{
    public string Name { get; set; } = string.Empty;
    public string Empty { get; set; } = string.Empty;
}
