using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for broader encoding support.
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Paths for the template, HTML snippet, and the generated report.
        string templatePath = "Template.docx";
        string htmlSnippetPath = "snippet.html";
        string outputPath = "Report.docx";

        // -----------------------------------------------------------------
        // Step 1: Create a simple template containing an HTML switch tag.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        // The tag will be replaced with the HTML content from the data source.
        builder.Writeln("<<[model.HtmlSnippet] -html>>");
        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Step 2: Prepare the HTML snippet file.
        // -----------------------------------------------------------------
        // For demonstration, write a sample HTML fragment if the file does not exist.
        if (!File.Exists(htmlSnippetPath))
        {
            const string sampleHtml = "<h2 style='color:steelblue;'>Hello from HTML snippet!</h2>" +
                                      "<p>This paragraph is inserted via LINQ Reporting.</p>";
            File.WriteAllText(htmlSnippetPath, sampleHtml);
        }

        // Read the HTML content from the external file.
        string htmlContent = File.ReadAllText(htmlSnippetPath);

        // -----------------------------------------------------------------
        // Step 3: Load the template and build the report.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);
        var model = new ReportModel { HtmlSnippet = htmlContent };

        ReportingEngine engine = new ReportingEngine();
        // The root object name must match the tag's reference ("model").
        engine.BuildReport(reportDoc, model, "model");

        // -----------------------------------------------------------------
        // Step 4: Save the generated report.
        // -----------------------------------------------------------------
        reportDoc.Save(outputPath);
    }
}

// Data model used by the LINQ Reporting engine.
public class ReportModel
{
    // Holds the HTML snippet that will be inserted into the document.
    public string HtmlSnippet { get; set; } = string.Empty;
}
