// ALL ATTEMPTS FAILED. Below is the last generated code.

using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;
using Aspose.Words.Saving.MultiPageLayout;

class Program
{
    static void Main()
    {
        // Path to the DOCX template that contains Aspose.Words reporting tags.
        string templatePath = @"C:\Data\Template.docx";

        // Path where the resulting multi‑page TIFF will be saved.
        string outputPath = @"C:\Data\Result.tiff";

        // Load the template document.
        Document doc = new Document(templatePath);

        // Prepare the data source that will be merged into the template.
        // This can be any object supported by ReportingEngine (e.g., a POCO, DataSet, etc.).
        var dataSource = new
        {
            Title = "Quarterly Report",
            Date = DateTime.Now,
            Items = new[]
            {
                new { Name = "Item A", Quantity = 10, Price = 9.99 },
                new { Name = "Item B", Quantity = 5,  Price = 19.95 },
                new { Name = "Item C", Quantity = 2,  Price = 99.00 }
            }
        };

        // Populate the template with the data.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, dataSource);

        // Configure image save options for TIFF output.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Render all pages into a single multi‑frame TIFF.
            MultiPageLayout = MultiPageLayout.TiffFrames(),

            // Optional: set resolution and image size if required.
            Resolution = 300,
            // ImageSize = new System.Drawing.Size(2480, 3508) // A4 at 300 dpi, uncomment if needed.
        };

        // Save the populated document as a multi‑page TIFF.
        doc.Save(outputPath, saveOptions);
    }
}
