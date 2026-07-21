using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some Aspose.Words features).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create a simple data model.
        var model = new ReportModel
        {
            Name = "John Doe",
            // This property will be empty after the report is built, leaving its paragraph empty.
            EmptyTag = string.Empty
        };

        // Build the template document programmatically.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Paragraph that will contain a real value.
        builder.Writeln("Customer: <<[model.Name]>>");

        // Paragraph that contains only a tag which resolves to an empty string.
        builder.Writeln("<<[model.EmptyTag]>>");

        // Configure the reporting engine to remove empty paragraphs after processing.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

        // Build the report using the model as the root data source.
        engine.BuildReport(doc, model, "model");

        // Save the resulting document.
        doc.Save("ReportOutput.docx");
    }
}

// Public data model class required by the LINQ Reporting engine.
public class ReportModel
{
    public string Name { get; set; } = string.Empty;
    public string EmptyTag { get; set; } = string.Empty;
}
