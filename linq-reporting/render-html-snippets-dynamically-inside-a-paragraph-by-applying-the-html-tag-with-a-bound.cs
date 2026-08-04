using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Model
{
    // HTML snippet to be rendered inside the paragraph.
    public string HtmlSnippet { get; set; } = "<b>Default Bold</b> and <i>Default Italic</i>";
}

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some environments).
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // ---------- Create the template document ----------
        const string templatePath = "template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Insert a paragraph containing the LINQ Reporting tag that renders HTML.
        // The tag uses the -html switch to treat the bound string as HTML.
        builder.Writeln("<<[model.HtmlSnippet] -html>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // ---------- Load the template and build the report ----------
        var reportDoc = new Document(templatePath);

        // Prepare the data model with a dynamic HTML snippet.
        var model = new Model
        {
            HtmlSnippet = "<p style='color:blue;'>This is <span style='font-weight:bold;'>dynamic</span> HTML content.</p>"
        };

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated document.
        const string outputPath = "output.docx";
        reportDoc.Save(outputPath);
    }
}
