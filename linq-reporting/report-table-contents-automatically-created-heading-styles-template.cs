using System;
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
            // Load the template document (create/load lifecycle rule).
            Document doc = new Document(templatePath);

            // Populate the template with the provided data using the reporting engine
            // (feature rule: BuildReport).
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, data);

            // Insert a Table of Contents at the beginning of the document.
            // The switches configure the TOC to include heading levels 1‑3,
            // make entries hyperlinks, and use the default formatting.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveToDocumentStart();
            builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
            builder.InsertBreak(BreakType.PageBreak); // Optional: start content on a new page.

            // Update all fields (including the TOC) so that page numbers are correct.
            doc.UpdateFields();

            // Save the final report (save lifecycle rule).
            doc.Save(outputPath);
        }
    }

    // Example usage.
    class Program
    {
        static void Main()
        {
            // Paths to the template and the output document.
            string templatePath = @"C:\Templates\ReportTemplate.docx";
            string outputPath = @"C:\Reports\GeneratedReport.docx";

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

            Console.WriteLine("Report generated successfully.");
        }
    }
}
