using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data model that mimics a database field containing HTML.
        ReportModel model = new()
        {
            HtmlContent = "<h1>Report Title</h1><p>This is a <b>bold</b> paragraph with <i>italic</i> text.</p>"
        };

        // Create a template document programmatically.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);
        builder.Writeln("Below is the HTML content from the data source:");
        // Insert the HTML expression tag using the -html switch.
        builder.Writeln("<<[model.HtmlContent] -html>>");
        // Save the template to disk.
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        Document reportDoc = new(templatePath);

        // Build the report using LINQ Reporting Engine.
        ReportingEngine engine = new();
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        reportDoc.Save(outputPath);
    }
}

// Public data model class.
public class ReportModel
{
    // HTML snippet that will be inserted into the document.
    public string HtmlContent { get; set; } = string.Empty;
}
