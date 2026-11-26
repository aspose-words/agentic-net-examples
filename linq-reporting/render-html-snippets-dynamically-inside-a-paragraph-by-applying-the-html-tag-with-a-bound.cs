using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // HTML snippet to be inserted into the document.
    public string HtmlSnippet { get; set; } = "<b>Bold Text</b> and <i>Italic Text</i>";
}

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some Aspose.Words features).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data.
        ReportModel model = new();

        // Create a template document programmatically.
        Document template = new();
        DocumentBuilder builder = new(template);
        builder.Writeln("Paragraph with HTML: <<[model.HtmlSnippet] -html>>");

        // Save the template to disk.
        string templatePath = "template.docx";
        template.Save(templatePath);

        // Load the template back for reporting.
        Document doc = new(templatePath);

        // Build the report using LINQ Reporting Engine.
        ReportingEngine engine = new();
        engine.BuildReport(doc, model, "model");

        // Save the generated document.
        string outputPath = "output.docx";
        doc.Save(outputPath);
    }
}
