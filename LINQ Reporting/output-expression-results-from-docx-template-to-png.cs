using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsTemplateToPng
{
    // Simple data source class used in the template.
    public class ReportData
    {
        public string Title { get; set; }
        public string Description { get; set; }
        public double Amount { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Path to the DOCX template that contains Aspose.Words reporting tags,
            // e.g. <<[data.Title]>>, <<[data.Description]>>, <<[data.Amount]:currency>>
            string templatePath = @"C:\Templates\ReportTemplate.docx";

            // Load the template document.
            Document doc = new Document(templatePath);

            // Create a data source instance with values that will replace the tags.
            var data = new ReportData
            {
                Title = "Quarterly Sales Report",
                Description = "Summary of sales performance for Q1 2024.",
                Amount = 1234567.89
            };

            // Build the report by merging the template with the data source.
            ReportingEngine engine = new ReportingEngine();
            // The name "data" is used inside the template to reference the object.
            engine.BuildReport(doc, data, "data");

            // Configure image save options for PNG output.
            ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
            {
                // Render the first page only (zero‑based index). Adjust if you need all pages.
                PageSet = new PageSet(0),

                // Optional: set resolution (dpi) for higher quality.
                Resolution = 300,

                // Optional: make background transparent.
                PaperColor = System.Drawing.Color.Transparent
            };

            // Output file path.
            string outputPath = @"C:\Output\ReportPage1.png";

            // Save the rendered page as PNG.
            doc.Save(outputPath, pngOptions);

            Console.WriteLine($"Report rendered to PNG: {outputPath}");
        }
    }
}
