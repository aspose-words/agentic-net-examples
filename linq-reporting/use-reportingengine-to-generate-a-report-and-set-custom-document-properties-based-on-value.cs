using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    public string Title { get; set; } = string.Empty;
    public string Author { get; set; } = string.Empty;
}

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        ReportModel model = new ReportModel
        {
            Title = "Quarterly Report",
            Author = "John Doe"
        };

        // Create a template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("Report: <<[model.Title]>>");
        builder.Writeln("Prepared by: <<[model.Author]>>");

        // Save the template to disk.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        template.Save(templatePath);

        // Load the template for reporting.
        Document report = new Document(templatePath);

        // Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(report, model, "model");

        // Set custom document properties based on the model values.
        if (report.CustomDocumentProperties["ReportTitle"] != null)
            report.CustomDocumentProperties["ReportTitle"].Value = model.Title;
        else
            report.CustomDocumentProperties.Add("ReportTitle", model.Title);

        if (report.CustomDocumentProperties["ReportAuthor"] != null)
            report.CustomDocumentProperties["ReportAuthor"].Value = model.Author;
        else
            report.CustomDocumentProperties.Add("ReportAuthor", model.Author);

        // Save the final report.
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        report.Save(reportPath);
    }
}
