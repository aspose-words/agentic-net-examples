using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsExample
{
    // Simple data source class – can be any POCO that matches the template tags.
    public class ReportData
    {
        public string Title { get; set; }
        public string Author { get; set; }
        public string Content { get; set; }
    }

    public class TemplateToJpegConverter
    {
        /// <summary>
        /// Loads a DOCX template, fills it with data using ReportingEngine,
        /// and saves each rendered page as a separate JPEG file.
        /// </summary>
        /// <param name="templatePath">Full path to the DOCX template.</param>
        /// <param name="data">Object containing data for the template.</param>
        /// <param name="outputFolder">Folder where JPEG files will be written.</param>
        public static void Convert(string templatePath, ReportData data, string outputFolder)
        {
            // Ensure the output directory exists.
            Directory.CreateDirectory(outputFolder);

            // Load the template document.
            Document doc = new Document(templatePath);

            // Populate the template with the provided data.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, data);

            // Prepare image save options for JPEG format.
            ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
            {
                // Optional: set higher quality (0‑100). 100 = best quality.
                JpegQuality = 100,
                // Optional: set resolution (dpi) if needed.
                Resolution = 300
            };

            // Render each page of the document to a separate JPEG file.
            for (int pageIndex = 0; pageIndex < doc.PageCount; pageIndex++)
            {
                // PageSet selects a single page by zero‑based index.
                jpegOptions.PageSet = new PageSet(pageIndex);

                // Build the output file name, e.g., Page_1.jpg, Page_2.jpg, …
                string outputPath = Path.Combine(outputFolder, $"Page_{pageIndex + 1}.jpg");

                // Save the current page as JPEG.
                doc.Save(outputPath, jpegOptions);
            }
        }

        // Example usage.
        public static void Main()
        {
            string templateFile = @"C:\Templates\ReportTemplate.docx";
            string outputDir = @"C:\Output\ReportPages";

            // Create a data object that matches the template placeholders.
            ReportData data = new ReportData
            {
                Title = "Quarterly Sales Report",
                Author = "John Doe",
                Content = "The sales increased by 15% compared to the previous quarter..."
            };

            // Perform the conversion.
            Convert(templateFile, data, outputDir);

            Console.WriteLine("Conversion completed.");
        }
    }
}
