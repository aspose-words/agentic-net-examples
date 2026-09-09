using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class HyperlinkModel
{
    // Original target object (Uri)
    public Uri Url { get; set; } = new Uri("https://example.com");

    // Text that will be displayed for the hyperlink
    public string DisplayText { get; set; } = "Example Site";

    // Convert the Uri to string before the engine processes it
    public string UrlString => Url.ToString();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some environments)
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // ---------- Create the template ----------
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Insert a link tag that uses the string representation of the Uri
        builder.Writeln("<<link [model.UrlString] [model.DisplayText]>>");

        // Save the template locally
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        template.Save(templatePath);

        // ---------- Load the template and build the report ----------
        var doc = new Document(templatePath);
        var model = new HyperlinkModel();

        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        doc.Save(outputPath);
    }
}
