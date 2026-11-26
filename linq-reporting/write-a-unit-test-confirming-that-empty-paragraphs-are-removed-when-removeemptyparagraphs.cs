using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingRemoveEmptyParagraphsTest
{
    // Simple data model used as the root object for the report.
    public class ReportModel
    {
        public string Title { get; set; } = "Sample Report";
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare file paths.
            string templatePath = "template.docx";
            string outputPath = "output.docx";

            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Paragraph with static text.
            builder.Writeln("=== Report Start ===");

            // Empty paragraph that should be removed.
            builder.Writeln(string.Empty);

            // LINQ Reporting tag that will be replaced with model data.
            builder.Writeln("Title: <<[model.Title]>>");

            // Another empty paragraph.
            builder.Writeln();

            // Paragraph with static text.
            builder.Writeln("=== Report End ===");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and build the report with RemoveEmptyParagraphs enabled.
            // -----------------------------------------------------------------
            Document docToReport = new Document(templatePath);
            ReportModel model = new ReportModel();

            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

            bool success = engine.BuildReport(docToReport, model, "model");
            if (!success)
            {
                Console.WriteLine("Report generation failed.");
                return;
            }

            // Save the generated report.
            docToReport.Save(outputPath);

            // -----------------------------------------------------------------
            // 3. Verify that no empty paragraphs remain in the generated document.
            // -----------------------------------------------------------------
            Document resultDoc = new Document(outputPath);
            var paragraphs = resultDoc.GetChildNodes(NodeType.Paragraph, true)
                                     .Cast<Paragraph>()
                                     .ToList();

            int emptyParagraphCount = paragraphs.Count(p => string.IsNullOrWhiteSpace(p.GetText()));

            if (emptyParagraphCount == 0)
            {
                Console.WriteLine("Test passed: No empty paragraphs remain.");
            }
            else
            {
                Console.WriteLine($"Test failed: {emptyParagraphCount} empty paragraph(s) found.");
            }
        }
    }
}
