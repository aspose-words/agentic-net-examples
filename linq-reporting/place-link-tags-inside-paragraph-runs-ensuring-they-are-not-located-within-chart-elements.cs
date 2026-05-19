using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts; // ChartType enum resides here
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "LinkTemplate.docx";
        const string reportPath = "LinkReport.docx";

        // -----------------------------------------------------------------
        // 1. Create the template document programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Normal paragraph with a LINQ Reporting link tag.
        builder.Writeln("Report with a hyperlink:");
        // The link tag uses the model's Url and LinkText properties.
        builder.Writeln("<<link [model.Url] [model.LinkText]>>");

        // Insert a chart to demonstrate that the link tag is NOT placed inside it.
        builder.Writeln();
        builder.InsertChart(ChartType.Column, 400, 300);
        builder.Writeln(); // Ensure the chart is closed with a paragraph break.

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // Sample data model for the report.
        ReportModel model = new ReportModel
        {
            Url = "https://example.com",
            LinkText = "Example Site"
        };

        // Configure and execute the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        reportDoc.Save(reportPath);
    }
}

// Simple data model used by the LINQ Reporting template.
public class ReportModel
{
    public string Url { get; set; } = "";
    public string LinkText { get; set; } = "";
}
