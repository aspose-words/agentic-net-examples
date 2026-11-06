using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    public string Url { get; set; } = "https://example.com";
    public string LinkText { get; set; } = "Visit Example";
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for any encoding needs.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for template and output documents.
        string templatePath = "template.docx";
        string outputPath = "output.docx";

        // -------------------------------------------------
        // Create the template document with a LINQ Reporting link tag.
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Dynamic hyperlink example:");
        // The link tag uses expressions that will be replaced from the model.
        builder.Writeln("<<link [model.Url] [model.LinkText]>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // Load the template and build the report.
        // -------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // Sample data model.
        ReportModel model = new ReportModel
        {
            Url = "https://www.aspose.com",
            LinkText = "Aspose.Words"
        };

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        reportDoc.Save(outputPath);
    }
}
