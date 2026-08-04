using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model used by the LINQ Reporting template.
    public class ReportModel
    {
        // Name will always have a value.
        public string Name { get; set; } = string.Empty;

        // Optional may be null or empty; when empty the paragraph should disappear.
        public string? Optional { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare folders.
            string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
            Directory.CreateDirectory(workDir);

            // -----------------------------------------------------------------
            // 1. Create a template document programmatically.
            // -----------------------------------------------------------------
            string templatePath = Path.Combine(workDir, "Template.docx");
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            builder.Writeln("=== Report Start ===");
            // This line will always be filled.
            builder.Writeln("Customer: <<[model.Name]>>");
            // This line may become empty after rendering.
            builder.Writeln("Notes: <<[model.Optional]>>");
            builder.Writeln("=== Report End ===");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and build the report.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);

            // Sample data where Optional is null – the paragraph should be removed.
            ReportModel data = new ReportModel
            {
                Name = "John Doe",
                Optional = null
            };

            // Configure the reporting engine to remove empty paragraphs.
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

            // Build the report. The root object name must match the tags in the template.
            engine.BuildReport(reportDoc, data, "model");

            // -----------------------------------------------------------------
            // 3. Save the final document.
            // -----------------------------------------------------------------
            string resultPath = Path.Combine(workDir, "Result.docx");
            reportDoc.Save(resultPath);

            // Indicate completion (no interactive prompts).
            Console.WriteLine($"Report generated: {resultPath}");
        }
    }
}
