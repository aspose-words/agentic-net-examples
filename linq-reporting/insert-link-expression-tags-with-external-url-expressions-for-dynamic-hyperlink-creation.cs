using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    public string Url { get; set; } = string.Empty;
    public string LinkText { get; set; } = string.Empty;
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create a template document with a LINQ Reporting link tag.
        var template = new Document();
        var builder = new DocumentBuilder(template);
        builder.Writeln("Visit our site: <<link [model.Url] [model.LinkText]>>");
        template.Save("Template.docx");

        // Load the template for reporting.
        var doc = new Document("Template.docx");

        // Prepare sample data.
        var model = new ReportModel
        {
            Url = "https://example.com",
            LinkText = "Example Site"
        };

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}
