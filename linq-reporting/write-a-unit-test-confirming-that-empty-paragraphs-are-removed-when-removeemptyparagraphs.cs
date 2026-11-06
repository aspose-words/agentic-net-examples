using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words if needed.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create a simple template document.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Header paragraph.
        builder.Writeln("Header");

        // Conditional paragraph that will be empty when IncludeDetail is false.
        builder.Writeln("<<if [model.IncludeDetail]>>Detail <<[model.Detail]>> <</if>>");

        // Footer paragraph.
        builder.Writeln("Footer");

        // Save the template (optional, for inspection).
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "template.docx");
        template.Save(templatePath);

        // Prepare the data model with IncludeDetail = false.
        ReportModel model = new()
        {
            IncludeDetail = false,
            Detail = "Some detail"
        };

        // Configure the reporting engine to remove empty paragraphs.
        ReportingEngine engine = new();
        engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

        // Build the report.
        bool success = engine.BuildReport(template, model, "model");

        // Verify that empty paragraphs were removed.
        // Expected paragraphs: Header and Footer only (2 paragraphs).
        int expectedParagraphCount = 2;
        int actualParagraphCount = template.FirstSection.Body.Paragraphs.Count;

        if (success && actualParagraphCount == expectedParagraphCount)
        {
            Console.WriteLine("Test passed: Empty paragraphs were removed.");
        }
        else
        {
            Console.WriteLine($"Test failed: Expected {expectedParagraphCount} paragraphs, but found {actualParagraphCount}.");
        }

        // Save the resulting document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");
        template.Save(outputPath);
    }

    public class ReportModel
    {
        public bool IncludeDetail { get; set; } = false;
        public string Detail { get; set; } = string.Empty;
    }
}
