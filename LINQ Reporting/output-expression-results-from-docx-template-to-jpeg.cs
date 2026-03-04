using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsTemplateToJpeg
{
    // Simple data source class used in the template.
    public class ReportData
    {
        public string Title { get; set; }
        public string Author { get; set; }
        public DateTime Date { get; set; }
        public decimal Amount { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Path to the DOCX template that contains Aspose.Words reporting tags,
            // e.g. <<[data.Title]>>, <<[data.Author]>>, etc.
            string templatePath = @"C:\Templates\ReportTemplate.docx";

            // Load the template document.
            Document doc = new Document(templatePath);

            // Prepare the data source instance.
            ReportData data = new ReportData
            {
                Title = "Quarterly Sales Report",
                Author = "John Doe",
                Date = DateTime.Today,
                Amount = 123456.78m
            };

            // Build the report by populating the template with the data.
            ReportingEngine engine = new ReportingEngine();
            // The second overload allows referencing the data source object itself via the name "data".
            engine.BuildReport(doc, data, "data");

            // Configure image save options for JPEG output.
            ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
            {
                // Optional: set JPEG quality (0‑100). Higher value = better quality, larger file.
                JpegQuality = 90,
                // Optional: enable high‑quality rendering for better visual fidelity.
                UseHighQualityRendering = true,
                // Optional: enable anti‑aliasing to smooth edges.
                UseAntiAliasing = true,
                // Optional: render only the first page (default behavior for image formats).
                // If you need a specific page, set the PageSet property, e.g.:
                // PageSet = new PageSet(0) // zero‑based index for the first page.
            };

            // Path for the resulting JPEG image.
            string outputPath = @"C:\Output\ReportImage.jpg";

            // Save the populated document as a JPEG image.
            doc.Save(outputPath, jpegOptions);

            Console.WriteLine("Report rendered to JPEG successfully.");
        }
    }
}
