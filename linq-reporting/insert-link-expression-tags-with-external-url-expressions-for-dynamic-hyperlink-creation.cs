using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    public string Url { get; set; } = "";
    public string LinkText { get; set; } = "";
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for full encoding support.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for template and output documents.
        string templatePath = "Template.docx";
        string outputPath = "Report.docx";

        // -------------------------------------------------
        // Create the LINQ Reporting template programmatically.
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a LINQ Reporting link tag that will be replaced with a dynamic hyperlink.
        // The tag uses the model's Url and LinkText properties.
        builder.Writeln("<<link [model.Url] [model.LinkText]>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // Load the template for report generation.
        // -------------------------------------------------
        Document loadedTemplate = new Document(templatePath);

        // Prepare the data model with sample values.
        ReportModel model = new ReportModel
        {
            Url = "https://www.example.com",
            LinkText = "Visit Example"
        };

        // Create the reporting engine and build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(loadedTemplate, model, "model");

        // Save the generated report.
        loadedTemplate.Save(outputPath);
    }
}
