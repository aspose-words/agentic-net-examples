using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsExample
{
    // Sample data source class – replace with your actual data model.
    public class ReportData
    {
        public string Title { get; set; }
        public string Author { get; set; }
        public DateTime Date { get; set; }
        // Add other properties referenced in the DOCX template.
    }

    class Program
    {
        static void Main()
        {
            // Path to the DOCX template that contains Aspose.Words reporting tags.
            string templatePath = @"C:\Templates\ReportTemplate.docx";

            // Load the template document.
            Document doc = new Document(templatePath);

            // Prepare the data source that will populate the template.
            ReportData data = new ReportData
            {
                Title = "Quarterly Sales Report",
                Author = "John Doe",
                Date = DateTime.Today
                // Populate additional fields as needed.
            };

            // Execute the reporting engine to merge data into the template.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, data);

            // Configure image save options for JPEG output.
            ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
            {
                // Optional: set image quality (0‑100). Higher = better quality, larger file.
                JpegQuality = 90,
                // Optional: enable high‑quality rendering for sharper results.
                UseHighQualityRendering = true,
                // Optional: enable anti‑aliasing.
                UseAntiAliasing = true
            };

            // Render each page of the populated document to a separate JPEG file.
            for (int pageIndex = 0; pageIndex < doc.PageCount; pageIndex++)
            {
                // Select the current page for rendering.
                jpegOptions.PageSet = new PageSet(pageIndex);

                // Build the output file name (e.g., Report_1.jpg, Report_2.jpg, ...).
                string outputPath = $@"C:\Output\Report_{pageIndex + 1}.jpg";

                // Save the selected page as a JPEG image.
                doc.Save(outputPath, jpegOptions);
            }

            // All pages have been saved as JPEG images.
            Console.WriteLine("Document rendered to JPEG successfully.");
        }
    }
}
