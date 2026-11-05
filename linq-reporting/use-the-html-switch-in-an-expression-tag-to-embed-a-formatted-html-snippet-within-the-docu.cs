using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for potential encoding needs.
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Create a template document with a LINQ Reporting HTML expression tag.
        const string templatePath = "template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Report with HTML snippet:");
        builder.Writeln("<<[model.HtmlSnippet] -html>>");
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        var doc = new Document(templatePath);

        // Prepare the data model containing the HTML snippet.
        var model = new ReportModel
        {
            HtmlSnippet = "<b>Hello <i>World</i></b><br/><p style='color:red;'>Red text</p>"
        };

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);
    }
}

// Data model used by the LINQ Reporting template.
public class ReportModel
{
    public string HtmlSnippet { get; set; } = string.Empty;
}
