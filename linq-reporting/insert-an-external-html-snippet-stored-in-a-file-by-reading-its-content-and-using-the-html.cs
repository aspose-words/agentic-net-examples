using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required by Aspose.Words for some encodings)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Define file paths
        string workingDir = Directory.GetCurrentDirectory();
        string templatePath = Path.Combine(workingDir, "template.docx");
        string htmlSnippetPath = Path.Combine(workingDir, "snippet.html");
        string outputPath = Path.Combine(workingDir, "output.docx");

        // Create a simple HTML snippet file
        File.WriteAllText(htmlSnippetPath,
            "<h1>Welcome to Aspose.Words</h1><p>This is an <b>HTML</b> snippet inserted via LINQ Reporting.</p>");

        // Create the LINQ Reporting template programmatically
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        // Insert the HTML switch tag that will be replaced by the snippet content
        builder.Writeln("<<[model.HtmlSnippet] -html>>");
        // Save the template to disk
        templateDoc.Save(templatePath);

        // Load the template for report generation
        Document reportDoc = new Document(templatePath);

        // Read the HTML snippet content
        string htmlContent = File.ReadAllText(htmlSnippetPath);

        // Prepare the data model
        ReportModel model = new()
        {
            HtmlSnippet = htmlContent
        };

        // Build the report using the LINQ Reporting engine
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated document
        reportDoc.Save(outputPath);
    }
}

// Data model used by the LINQ Reporting template
public class ReportModel
{
    public string HtmlSnippet { get; set; } = string.Empty;
}
