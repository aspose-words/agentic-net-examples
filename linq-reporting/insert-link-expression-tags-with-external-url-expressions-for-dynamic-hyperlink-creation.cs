using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // URL to link to
    public string Url { get; set; } = string.Empty;
    // Text displayed for the hyperlink
    public string LinkText { get; set; } = string.Empty;
}

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report
        string templatePath = "template.docx";
        string outputPath = "output.docx";

        // -------------------------------------------------
        // Create the template document with a LINQ Reporting link tag
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Dynamic hyperlink example:");
        // The link tag uses expressions that will be replaced at runtime
        builder.Writeln("<<link [model.Url] [model.LinkText]>>");

        // Save the template to disk
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // Load the template and build the report
        // -------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // Sample data model
        ReportModel model = new ReportModel
        {
            Url = "https://www.example.com",
            LinkText = "Visit Example"
        };

        // Create the reporting engine and generate the report
        ReportingEngine engine = new ReportingEngine();
        bool success = engine.BuildReport(reportDoc, model, "model");

        // Save the generated report
        reportDoc.Save(outputPath);

        // Indicate completion (no interactive prompts)
        Console.WriteLine($"Report generation {(success ? "succeeded" : "failed")}. Output saved to '{outputPath}'.");
    }
}
