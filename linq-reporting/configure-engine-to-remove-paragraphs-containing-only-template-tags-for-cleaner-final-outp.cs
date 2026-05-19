using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model used by the LINQ Reporting template.
    public class ReportModel
    {
        // Required property – initialized to avoid nullable warnings.
        public string Name { get; set; } = string.Empty;

        // Optional property – can be null; when null the corresponding paragraph should disappear.
        public string? Optional { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the temporary template and the final report.
            string templatePath = Path.Combine(Environment.CurrentDirectory, "Template.docx");
            string outputPath   = Path.Combine(Environment.CurrentDirectory, "Report.docx");

            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Paragraph that will always contain data.
            builder.Writeln("Name: <<[model.Name]>>");

            // Paragraph that contains only a tag. If the tag resolves to an empty value,
            // the whole paragraph should be removed by the reporting engine.
            builder.Writeln("<<[model.Optional]>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template for report generation.
            // -----------------------------------------------------------------
            Document doc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Prepare the data source.
            // -----------------------------------------------------------------
            ReportModel model = new ReportModel
            {
                Name = "John Doe",
                Optional = null // This will cause the second paragraph to become empty.
            };

            // -----------------------------------------------------------------
            // 4. Configure and run the LINQ Reporting engine.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine
            {
                // Remove paragraphs that become empty after tag processing.
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };

            // Build the report. The root object name must match the tag prefix ("model").
            engine.BuildReport(doc, model, "model");

            // -----------------------------------------------------------------
            // 5. Save the final report.
            // -----------------------------------------------------------------
            doc.Save(outputPath);

            // Inform the user (no interactive input required).
            Console.WriteLine($"Report generated successfully: {outputPath}");
        }
    }
}
