using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words compatibility.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare a simple data model with URL and display text.
        var model = new ReportModel
        {
            Url = "https://example.com",
            LinkText = "Visit Example"
        };

        // Create the template document programmatically.
        var templatePath = "Template.docx";
        var builder = new DocumentBuilder();
        // Insert a LINQ Reporting link tag that uses the model fields.
        builder.Writeln("<<link [model.Url] [model.LinkText]>>");
        builder.Document.Save(templatePath);

        // Load the template for report generation.
        var doc = new Document(templatePath);

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        var outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}

// Public data model class required by the template.
public class ReportModel
{
    public string Url { get; set; } = string.Empty;
    public string LinkText { get; set; } = string.Empty;
}
