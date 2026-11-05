using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class ReportModel
{
    public string Url { get; set; } = "https://example.com";
    public string LinkText { get; set; } = "Example Site";
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data.
        ReportModel model = new();

        // Create a template document programmatically.
        string templatePath = "Template.docx";
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);

        // Write a paragraph with a link tag inside a run.
        builder.Writeln("Visit our website:");
        builder.Write("<<link [model.Url] [model.LinkText]>>");
        builder.Writeln(); // End the paragraph.

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new(templatePath);

        // Build the report using LINQ Reporting Engine.
        ReportingEngine engine = new();
        bool success = engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        string outputPath = "Report.docx";
        reportDoc.Save(outputPath);

        // Optionally, indicate success (no interactive output required).
        // The program ends here.
    }
}
