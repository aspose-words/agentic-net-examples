using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Model
{
    // HTML snippet that will be rendered inside the paragraph.
    public string HtmlSnippet { get; set; } = "<b>Bold Text</b> and <i>Italic Text</i>";
}

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create the template document programmatically.
        // -----------------------------------------------------------------
        const string templateFile = "Template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Insert a paragraph containing the LINQ Reporting HTML tag.
        // The tag will be replaced with the HTML from Model.HtmlSnippet.
        builder.Writeln("<<[model.HtmlSnippet] -html>>");

        // Save the template to disk.
        templateDoc.Save(templateFile);

        // -----------------------------------------------------------------
        // 2. Load the template and prepare the data model.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templateFile);
        var data = new Model(); // Model instance with sample HTML.

        // -----------------------------------------------------------------
        // 3. Build the report using the ReportingEngine.
        // -----------------------------------------------------------------
        var engine = new ReportingEngine();
        // The root object name in the template is "model".
        engine.BuildReport(reportDoc, data, "model");

        // -----------------------------------------------------------------
        // 4. Save the generated report.
        // -----------------------------------------------------------------
        const string outputFile = "Report.docx";
        reportDoc.Save(outputFile);
    }
}
