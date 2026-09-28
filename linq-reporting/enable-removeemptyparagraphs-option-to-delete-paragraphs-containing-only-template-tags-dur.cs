using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Model
{
    public string Name { get; set; } = "John Doe";
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for any required encodings.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create sample data model.
        var model = new Model();

        // Create a new Word document and insert LINQ Reporting tags.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Paragraph that contains only a template tag – should be removed.
        builder.Writeln("<<[model.Name]>>");

        // Paragraph with static text – should remain.
        builder.Writeln("Report generated successfully.");

        // Initialize the reporting engine.
        var engine = new ReportingEngine();

        // Enable the option to remove empty paragraphs that contain only tags.
        engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

        // Build the report using the model as the root object named "model".
        engine.BuildReport(doc, model, "model");

        // Save the generated document.
        const string outputPath = "Report_Output.docx";
        doc.Save(outputPath);
    }
}
