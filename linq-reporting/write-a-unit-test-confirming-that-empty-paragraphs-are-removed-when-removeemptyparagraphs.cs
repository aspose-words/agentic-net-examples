using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Simple data model used by the LINQ Reporting template.
    public class Model
    {
        // This property is intentionally empty to generate an empty paragraph after processing.
        public string Empty { get; set; } = string.Empty;

        // Additional property to demonstrate normal data insertion.
        public string Text { get; set; } = "Hello";
    }

    public class Program
    {
        // Entry point of the console application.
        public static void Main()
        {
            // Prepare file paths.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
            Directory.CreateDirectory(outputDir);
            string templatePath = Path.Combine(outputDir, "template.docx");
            string resultPath = Path.Combine(outputDir, "result.docx");

            // -----------------------------------------------------------------
            // 1. Create a template document containing a tag that resolves to an empty string.
            // -----------------------------------------------------------------
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            builder.Writeln("Before");                     // First paragraph.
            builder.Writeln("<<[model.Empty]>>");          // This will become empty after the report is built.
            builder.Writeln("After");                      // Third paragraph.

            // Save the template to disk.
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and build the report with RemoveEmptyParagraphs enabled.
            // -----------------------------------------------------------------
            Document report = new Document(templatePath);
            Model data = new Model(); // Empty property already set.

            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.RemoveEmptyParagraphs; // Enable removal of empty paragraphs.
            engine.BuildReport(report, data, "model");

            // Save the generated report.
            report.Save(resultPath);

            // -----------------------------------------------------------------
            // 3. Verify that the empty paragraph has been removed.
            // -----------------------------------------------------------------
            // The expected text after removal: "Before\rAfter"
            string actualText = report.GetText().Trim(); // Trim removes leading/trailing whitespace.
            string expectedText = "Before\rAfter";

            if (actualText == expectedText)
                Console.WriteLine("Test passed: Empty paragraph was successfully removed.");
            else
                Console.WriteLine($"Test failed: Expected \"{expectedText}\", but got \"{actualText}\".");

            // Optional: display the location of the generated files.
            Console.WriteLine($"Template saved to: {templatePath}");
            Console.WriteLine($"Result saved to:   {resultPath}");
        }
    }
}
