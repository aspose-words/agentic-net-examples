using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsReportingToJpeg
{
    // Simple data source class for the LINQ Reporting Engine.
    public class ReportData
    {
        public string Title { get; set; }
        public int Quantity { get; set; }
        public decimal Price { get; set; }

        public ReportData(string title, int quantity, decimal price)
        {
            Title = title;
            Quantity = quantity;
            Price = price;
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the PDF template that contains LINQ Reporting tags, e.g. <<[Data.Title]>>.
            const string templatePath = @"C:\Templates\ReportTemplate.pdf";

            // Load the PDF template into an Aspose.Words Document.
            Document doc = new Document(templatePath);

            // Create an instance of the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();

            // Prepare the data source.
            ReportData data = new ReportData("Sample Product", 5, 19.99m);

            // Populate the template with data.
            // The second overload allows referencing the data source object itself via the name "Data".
            engine.BuildReport(doc, data, "Data");

            // Configure image save options for JPEG output.
            ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
            {
                // Render the first page only (zero‑based index).
                PageSet = new PageSet(0),

                // Set JPEG quality (0‑100). 90 gives high quality with reasonable size.
                JpegQuality = 90,

                // Optional: set resolution (dpi) if higher quality is needed.
                Resolution = 300
            };

            // Save the populated document as a JPEG image.
            const string outputPath = @"C:\Output\ReportResult.jpg";
            doc.Save(outputPath, jpegOptions);
        }
    }
}
