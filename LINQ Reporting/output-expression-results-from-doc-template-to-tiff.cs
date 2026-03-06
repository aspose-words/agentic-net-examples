using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the DOC template.
        Document doc = new Document("Template.docx");

        // Example data source that matches the template's expression tags.
        var dataSource = new
        {
            Name = "John Doe",
            Date = DateTime.Now,
            Amount = 1234.56
        };

        // Populate the template with the data.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, dataSource, "ds"); // template uses <<[ds.Name]>> etc.

        // Set up TIFF save options to create a multi‑frame TIFF (one frame per page).
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff);
        tiffOptions.PageLayout = MultiPageLayout.TiffFrames(); // each page as separate frame
        tiffOptions.Resolution = 300; // optional: set DPI

        // Save the rendered document as a TIFF file.
        doc.Save("Result.tiff", tiffOptions);
    }
}
