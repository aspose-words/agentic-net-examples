using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Category determines which HTML fragment file to load.
    public string Category { get; set; } = "A";

    // Returns the full path to the HTML fragment based on the Category value.
    public string HtmlFilePath => Path.Combine(Directory.GetCurrentDirectory(), $"{Category}.html");

    // Returns the content of the HTML fragment file.
    public string HtmlContent
    {
        get
        {
            // Ensure the file exists; if not, return an empty string to avoid null issues.
            return File.Exists(HtmlFilePath) ? File.ReadAllText(HtmlFilePath) : string.Empty;
        }
    }
}

public class Program
{
    public static void Main()
    {
        // Working directory.
        string workDir = Directory.GetCurrentDirectory();

        // Create sample HTML fragments.
        File.WriteAllText(Path.Combine(workDir, "A.html"), "<p><b>Category A Content</b></p>");
        File.WriteAllText(Path.Combine(workDir, "B.html"), "<p><i>Category B Content</i></p>");

        // -----------------------------------------------------------------
        // Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Report Header");
        // Insert the HTML fragment dynamically using the -html switch.
        builder.Writeln("<<[model.HtmlContent] -html>>");
        builder.Writeln("Report Footer");

        // Save the template to disk.
        string templatePath = Path.Combine(workDir, "template.docx");
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document loadedTemplate = new Document(templatePath);

        // Prepare the data model.
        ReportModel model = new()
        {
            Category = "B" // Change to "A" or "B" to select different fragments.
        };

        // Build the report.
        ReportingEngine engine = new();
        engine.BuildReport(loadedTemplate, model, "model");

        // Save the generated report.
        string outputPath = Path.Combine(workDir, "output.docx");
        loadedTemplate.Save(outputPath);
    }
}
