using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsExample
{
    // Simple data source class used in the template.
    public class ReportData
    {
        public string Title { get; set; }
        public string[] Items { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Path to the DOCX template that contains Aspose.Words reporting tags.
            string templatePath = @"C:\Templates\ReportTemplate.docx";

            // Load the template document.
            Document doc = new Document(templatePath);

            // Prepare data for the template.
            var data = new ReportData
            {
                Title = "Quarterly Sales Report",
                Items = new[] { "Item A", "Item B", "Item C" }
            };

            // Populate the template using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, data, "data"); // "data" is the name used in the template.

            // Render each page of the populated document to a separate PNG file.
            ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
            {
                // Optional: set resolution (dpi) for higher quality.
                Resolution = 300
            };

            // Loop through all pages in the document.
            for (int pageIndex = 0; pageIndex < doc.PageCount; pageIndex++)
            {
                // Select the current page.
                pngOptions.PageSet = new PageSet(pageIndex);

                // Build the output file name (e.g., Report_1.png, Report_2.png, ...).
                string outputPath = Path.Combine(
                    @"C:\Output",
                    $"Report_{pageIndex + 1}.png");

                // Save the current page as PNG.
                doc.Save(outputPath, pngOptions);
            }

            Console.WriteLine("Document rendered to PNG pages successfully.");
        }
    }
}
