using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class HtmlLinqReportingExample
{
    // Simple data model representing a database record with an HTML field.
    public class Record
    {
        // HTML content that will be inserted into the document.
        public string HtmlContent { get; set; } = string.Empty;
    }

    public static void Main()
    {
        // Prepare sample data (simulating a database field).
        var record = new Record
        {
            HtmlContent = "<b>Hello</b> <i>World</i>!<br/><span style=\"color:blue;\">Blue text</span>"
        };

        // Create a template document programmatically.
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Insert the LINQ Reporting tag that will render the HTML content.
        builder.Writeln("<<[model.HtmlContent] -html>>");

        // Save the template to disk.
        doc.Save(templatePath);

        // Load the template for report generation.
        var template = new Document(templatePath);

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(template, record, "model");

        // Save the generated document.
        var outputPath = "Report_Output.docx";
        template.Save(outputPath);

        // Optionally, indicate completion (no interactive prompts).
        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}
