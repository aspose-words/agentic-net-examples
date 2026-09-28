using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    public string Title { get; set; } = string.Empty;
}

public class Program
{
    public static void Main()
    {
        // Ensure the output folder exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a template document with empty paragraphs.
        string templatePath = Path.Combine(outputDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a data tag.
        builder.Writeln("<<[model.Title]>>");
        // Empty paragraph (should be removed).
        builder.Writeln("");
        // Paragraph with content.
        builder.Writeln("This is a content paragraph.");
        // Another empty paragraph.
        builder.Writeln("");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // Prepare the data model.
        ReportModel model = new ReportModel { Title = "Sample Report" };

        // Configure the reporting engine to remove empty paragraphs.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

        // Build the report.
        engine.BuildReport(reportDoc, model, "model");

        // Verify that there are no empty paragraphs.
        bool hasEmptyParagraph = false;
        foreach (Paragraph para in reportDoc.FirstSection.Body.Paragraphs)
        {
            // GetText includes the paragraph break; trim to check for content.
            if (string.IsNullOrWhiteSpace(para.GetText()))
            {
                hasEmptyParagraph = true;
                break;
            }
        }

        // Output the test result.
        Console.WriteLine(hasEmptyParagraph ? "Test failed: Empty paragraphs remain." : "Test passed: Empty paragraphs removed.");

        // Save the resulting document (optional).
        string resultPath = Path.Combine(outputDir, "result.docx");
        reportDoc.Save(resultPath);
    }
}
