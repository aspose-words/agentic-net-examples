using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Simple data model used by the LINQ Reporting engine.
    public class ReportModel
    {
        // This property will be empty, causing its paragraph to become empty after processing.
        public string EmptyValue { get; set; } = string.Empty;

        // Regular property that will appear in the output.
        public string Name { get; set; } = "John Doe";
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare folders.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // -----------------------------------------------------------------
            // 1. Create a template document containing a normal paragraph,
            //    an empty-value paragraph, and a trailing paragraph.
            // -----------------------------------------------------------------
            string templatePath = Path.Combine(outputDir, "Template.docx");
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Paragraph that will contain a real value.
            builder.Writeln("<<[model.Name]>>");

            // Paragraph that will be empty after the engine processes it.
            builder.Writeln("<<[model.EmptyValue]>>");

            // Final paragraph to verify the document still ends correctly.
            builder.Writeln("End");

            // Save the template.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template and build the report with RemoveEmptyParagraphs enabled.
            // -----------------------------------------------------------------
            Document doc = new Document(templatePath);
            ReportModel model = new ReportModel();

            ReportingEngine engine = new ReportingEngine
            {
                // Instruct the engine to delete paragraphs that become empty.
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };

            // Build the report. The root name must match the tag prefix used in the template.
            engine.BuildReport(doc, model, "model");

            // Save the generated document.
            string resultPath = Path.Combine(outputDir, "Result.docx");
            doc.Save(resultPath);

            // -----------------------------------------------------------------
            // 3. Verify that the empty paragraph was removed.
            //    The document should contain exactly two paragraphs: "John Doe" and "End".
            // -----------------------------------------------------------------
            int paragraphCount = doc.GetChildNodes(NodeType.Paragraph, true).Count;

            if (paragraphCount == 2)
            {
                Console.WriteLine("Test passed: Empty paragraph was removed.");
            }
            else
            {
                Console.WriteLine($"Test failed: Expected 2 paragraphs, but found {paragraphCount}.");
            }
        }
    }
}
