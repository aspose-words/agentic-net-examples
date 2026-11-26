using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Model
{
    // Property that will hold the HTML snippet read from an external file.
    public string HtmlSnippet { get; set; } = string.Empty;
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for any encoding needs.
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Define file paths in the current working directory.
        string workDir = Directory.GetCurrentDirectory();
        string templatePath = Path.Combine(workDir, "Template.docx");
        string htmlPath = Path.Combine(workDir, "snippet.html");
        string outputPath = Path.Combine(workDir, "ReportWithHtml.docx");

        // Create a sample HTML snippet file.
        File.WriteAllText(htmlPath, "<p style=\"color:blue;\">This is <b>HTML</b> inserted from external file.</p>");

        // Create the template document programmatically.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Report Title");
        // Insert the HTML switch that will be replaced by the model's HtmlSnippet property.
        builder.Writeln("<<[model.HtmlSnippet] -html>>");
        templateDoc.Save(templatePath);

        // Load the template document for reporting.
        Document doc = new Document(templatePath);

        // Prepare the data model, reading the HTML snippet from the file.
        Model model = new Model
        {
            HtmlSnippet = File.ReadAllText(htmlPath)
        };

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save(outputPath);
    }
}
