using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // HTML snippet that will be inserted into the document.
    public string HtmlSnippet { get; set; } = "<p style=\"color:blue;\">This is <b>bold</b> and <i>italic</i> text.</p>";
}

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "Template.docx";
        const string outputPath = "Report.docx";

        // -----------------------------------------------------------------
        // Create the template document with a LINQ Reporting HTML expression.
        // -----------------------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Insert the HTML expression tag that will render the HTML snippet.
        builder.Writeln("<<[model.HtmlSnippet] -html>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // Load the template and build the report with data.
        // -------------------------------------------------
        var doc = new Document(templatePath);
        var model = new ReportModel();

        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the final report.
        doc.Save(outputPath);
    }
}
