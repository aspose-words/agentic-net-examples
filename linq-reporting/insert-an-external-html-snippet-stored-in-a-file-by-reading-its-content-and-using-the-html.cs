using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Paths for the files used in the example.
        string workingDir = Directory.GetCurrentDirectory();
        string templatePath = Path.Combine(workingDir, "Template.docx");
        string htmlSnippetPath = Path.Combine(workingDir, "snippet.html");
        string outputPath = Path.Combine(workingDir, "Report.docx");

        // -----------------------------------------------------------------
        // 1. Create a simple HTML snippet file that will be inserted later.
        // -----------------------------------------------------------------
        const string sampleHtml = @"<h2 style='color:steelblue;'>Welcome to Aspose.Words</h2>
<p>This paragraph is <b>bold</b> and this one is <i>italic</i>.</p>";
        File.WriteAllText(htmlSnippetPath, sampleHtml);

        // --------------------------------------------------------------
        // 2. Build a LINQ Reporting template programmatically.
        // --------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Plain text title.
        builder.Writeln("=== LINQ Reporting Example ===");
        builder.Writeln();

        // Insert the LINQ Reporting tag that will render the HTML snippet.
        // The '-html' switch tells the engine to treat the expression as HTML.
        builder.Writeln("<<[model.HtmlSnippet] -html>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // --------------------------------------------------------------
        // 3. Load the HTML snippet from the external file.
        // --------------------------------------------------------------
        string htmlContent = File.ReadAllText(htmlSnippetPath);

        // --------------------------------------------------------------
        // 4. Prepare the data model for the report.
        // --------------------------------------------------------------
        var model = new ReportModel
        {
            HtmlSnippet = htmlContent
        };

        // --------------------------------------------------------------
        // 5. Load the template and build the report.
        // --------------------------------------------------------------
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();

        // BuildReport overload that includes the root name ("model").
        bool success = engine.BuildReport(reportDoc, model, "model");

        // Optional: check that the template was parsed without errors.
        if (!success)
        {
            Console.WriteLine("Warning: The template contained syntax errors.");
        }

        // --------------------------------------------------------------
        // 6. Save the generated report.
        // --------------------------------------------------------------
        reportDoc.Save(outputPath);
    }
}

// ----------------------------------------------------------------------
// Public data model used by the LINQ Reporting engine.
// ----------------------------------------------------------------------
public class ReportModel
{
    // The HTML snippet that will be inserted into the document.
    public string HtmlSnippet { get; set; } = string.Empty;
}
