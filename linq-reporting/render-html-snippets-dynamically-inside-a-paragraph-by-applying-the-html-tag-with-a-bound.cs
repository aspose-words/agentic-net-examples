using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // HTML snippet that will be inserted into the paragraph.
    public string HtmlSnippet { get; set; } = "<b>Default</b>";
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for any legacy encodings that Aspose.Words might need.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        const string templatePath = "Template.docx";
        const string outputPath = "Report.docx";

        // -------------------------------------------------
        // 1. Create the template document programmatically.
        // -------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Insert a paragraph containing the LINQ Reporting HTML tag.
        // The tag <<[model.HtmlSnippet] -html>> tells the engine to render the string as HTML.
        builder.Writeln("<<[model.HtmlSnippet] -html>>");

        // Save the template to disk before building the report.
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // 2. Load the template and prepare the data model.
        // -------------------------------------------------
        var reportDoc = new Document(templatePath);

        var model = new ReportModel
        {
            HtmlSnippet = "<span style='color:green;'>Dynamic <b>HTML</b> content.</span>"
        };

        // -------------------------------------------------
        // 3. Build the report using the LINQ Reporting engine.
        // -------------------------------------------------
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // -------------------------------------------------
        // 4. Save the final document.
        // -------------------------------------------------
        reportDoc.Save(outputPath);
    }
}
