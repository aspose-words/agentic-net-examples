using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for the template and the generated report.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        string reportPath = Path.Combine(outputDir, "Report.docx");

        // ---------- Create the template document ----------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a bookmark tag whose name is taken from the data model.
        // The expression [model.BookmarkName] will be evaluated during report building.
        builder.Writeln("<<bookmark [model.BookmarkName]>>");
        builder.Writeln("This text is inside the bookmark.");
        builder.Writeln("<</bookmark>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // ---------- Prepare the data model ----------
        ReportModel model = new ReportModel
        {
            // The bookmark name must be a non‑empty string.
            BookmarkName = "SampleBookmark"
        };

        // ---------- Load the template and build the report ----------
        Document docToReport = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();

        // BuildReport with the root object name "model" to match the tags.
        engine.BuildReport(docToReport, model, "model");

        // Save the final report.
        docToReport.Save(reportPath);
    }
}

// Simple data model used by the LINQ Reporting engine.
public class ReportModel
{
    // Non‑nullable property initialized to avoid warnings.
    public string BookmarkName { get; set; } = string.Empty;
}
