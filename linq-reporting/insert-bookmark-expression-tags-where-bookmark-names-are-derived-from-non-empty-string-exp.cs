using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using System.Text;

public class ReportModel
{
    public string Title { get; set; } = "Sample Title";
    public string BookmarkName { get; set; } = "MyBookmark";
}

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some environments)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create the template document with a bookmark tag whose name comes from a non‑empty expression
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("<<bookmark [model.BookmarkName]>>");
        builder.Writeln("<<[model.Title]>>");
        builder.Writeln("<</bookmark>>");

        doc.Save(templatePath);

        // Load the template for report generation
        var template = new Document(templatePath);

        // Prepare the data model
        var model = new ReportModel();

        // Build the report
        var engine = new ReportingEngine();
        engine.BuildReport(template, model, "model");

        // Save the generated document
        var outputPath = "Output.docx";
        template.Save(outputPath);
    }
}
