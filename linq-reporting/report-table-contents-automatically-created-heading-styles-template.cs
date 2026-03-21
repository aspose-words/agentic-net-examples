using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace ReportGeneration
{
    // Simple data source class used by the reporting engine.
    public class ReportData
    {
        public string Title { get; set; }
        public string[] Sections { get; set; }
    }

    public class ReportGenerator
    {
        // Generates a report from a template, inserts a TOC based on heading styles,
        // updates fields and saves the final document.
        public void GenerateReport(string templatePath, string outputPath, ReportData data)
        {
            // Load the template document.
            Document doc = new Document(templatePath);

            // Populate the template with the provided data using the reporting engine.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, data);

            // Insert a Table of Contents at the beginning of the document.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveToDocumentStart();
            builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
            builder.InsertBreak(BreakType.PageBreak); // Optional: start content on a new page.

            // Update all fields (including the TOC) so that page numbers are correct.
            doc.UpdateFields();

            // Ensure the output directory exists.
            Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);

            // Save the final report.
            doc.Save(outputPath);
        }
    }

    class Program
    {
        static void Main()
        {
            // Determine paths relative to the current directory.
            string baseDir = AppContext.BaseDirectory;
            string templatePath = Path.Combine(baseDir, "ReportTemplate.docx");
            string outputPath = Path.Combine(baseDir, "GeneratedReport.docx");

            // Create a simple template if it does not exist.
            if (!File.Exists(templatePath))
            {
                Document templateDoc = new Document();
                DocumentBuilder tb = new DocumentBuilder(templateDoc);

                // Title placeholder.
                tb.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
                tb.Write("{{Title}}");
                tb.InsertParagraph();

                // Section placeholders using a repeating region.
                tb.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
                tb.Write("{{#Sections}}");
                tb.InsertParagraph();
                tb.Write("{{.}}");
                tb.InsertParagraph();
                tb.Write("{{/Sections}}");
                tb.InsertParagraph();

                // Save the template.
                templateDoc.Save(templatePath);
            }

            // Sample data to populate the template.
            ReportData data = new ReportData
            {
                Title = "Quarterly Sales Report",
                Sections = new[]
                {
                    "Executive Summary",
                    "Market Analysis",
                    "Sales Figures",
                    "Future Outlook"
                }
            };

            // Generate the report.
            ReportGenerator generator = new ReportGenerator();
            generator.GenerateReport(templatePath, outputPath, data);

            Console.WriteLine($"Report generated successfully at: {outputPath}");
        }
    }
}
