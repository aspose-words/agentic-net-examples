using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Holds the HTML snippet read from an external file.
    public string HtmlSnippet { get; set; }

    public ReportModel(string htmlFilePath)
    {
        // Read the entire HTML file content.
        HtmlSnippet = File.ReadAllText(htmlFilePath);
    }
}

public class Program
{
    public static void Main()
    {
        // Define file paths relative to the current working directory.
        string workingDir = Directory.GetCurrentDirectory();
        string htmlFilePath = Path.Combine(workingDir, "snippet.html");
        string templatePath = Path.Combine(workingDir, "template.docx");
        string outputPath = Path.Combine(workingDir, "report.docx");

        // Create a simple HTML snippet file.
        File.WriteAllText(htmlFilePath,
            "<h2>Hello from HTML snippet</h2>" +
            "<p>This paragraph is inserted via LINQ Reporting.</p>");

        // Initialize the data model with the HTML content.
        ReportModel model = new ReportModel(htmlFilePath);

        // -----------------------------------------------------------------
        // Create the template document programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert the LINQ Reporting tag that will render the HTML.
        // The "-html" switch tells the engine to treat the expression as HTML.
        builder.Writeln("<<[model.HtmlSnippet] -html>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // Load the template back before building the report (required by the workflow).
        Document reportDoc = new Document(templatePath);

        // -----------------------------------------------------------------
        // Build the report using ReportingEngine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the final document containing the inserted HTML.
        reportDoc.Save(outputPath);
    }
}
