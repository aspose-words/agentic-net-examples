using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace RemoveEmptyParagraphsDemo
{
    // Simple data model used by the LINQ Reporting template.
    public class ReportModel
    {
        // This property will be empty, causing the paragraph that contains only its tag to be removed.
        public string EmptyTag { get; set; } = string.Empty;

        // Additional property to demonstrate that the report still contains content.
        public string Greeting { get; set; } = "Hello, Aspose.Words!";
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            string templatePath = Path.Combine(Environment.CurrentDirectory, "Template.docx");
            string reportPath   = Path.Combine(Environment.CurrentDirectory, "Report.docx");

            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Paragraph that contains only a tag which resolves to an empty string.
            builder.Writeln("<<[model.EmptyTag]>>");

            // Paragraph with a normal tag – this will remain in the final report.
            builder.Writeln("<<[model.Greeting]>>");

            // Save the template to disk before loading it for the report generation.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template document.
            // -----------------------------------------------------------------
            Document doc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Configure the ReportingEngine.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine
            {
                // Enable removal of paragraphs that become empty after tag processing.
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };

            // -----------------------------------------------------------------
            // 4. Build the report.
            // -----------------------------------------------------------------
            ReportModel model = new ReportModel(); // EmptyTag is empty, Greeting has a value.
            engine.BuildReport(doc, model, "model");

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            doc.Save(reportPath);

            // Optional: indicate completion (no interactive input required).
            Console.WriteLine("Report generated successfully at: " + reportPath);
        }
    }
}
