using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    public string Title { get; set; } = "Sample Section";
    public string BookmarkName { get; set; } = "SampleBookmark";
    public string LinkText { get; set; } = "go to section";
}

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for the template and the final report.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        string reportPath = Path.Combine(outputDir, "Report.docx");

        // ---------- Create template ----------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a bookmark whose name comes from the data model.
        builder.Writeln("<<bookmark [model.BookmarkName]>>");
        builder.Writeln("<<[model.Title]>>");
        builder.Writeln("<</bookmark>>");

        // Insert a hyperlink that points to the same bookmark.
        builder.Writeln("See <<link [model.BookmarkName] [model.LinkText]>> for details.");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // ---------- Load template and build report ----------
        Document reportDoc = new Document(templatePath);
        ReportModel model = new ReportModel(); // sample data

        ReportingEngine engine = new ReportingEngine();
        // No special options required for this scenario.
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        reportDoc.Save(reportPath);
    }
}
