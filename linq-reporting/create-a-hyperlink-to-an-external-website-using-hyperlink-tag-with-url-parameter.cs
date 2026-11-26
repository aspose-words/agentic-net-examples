using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    public string Url { get; set; } = "https://www.example.com";
    public string LinkText { get; set; } = "Example Website";
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words if needed.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare file paths.
        string templatePath = "Template.docx";
        string outputPath = "Report.docx";

        // Create the template document with a hyperlink tag.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Visit our site: <<link [model.Url] [model.LinkText]>>");
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // Prepare the data model.
        ReportModel model = new ReportModel();

        // Build the report using LINQ Reporting Engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        reportDoc.Save(outputPath);
    }
}
