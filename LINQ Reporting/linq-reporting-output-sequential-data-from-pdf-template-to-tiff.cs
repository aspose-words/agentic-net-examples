using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the PDF template document.
        Document doc = new Document("Template.pdf");

        // Prepare a simple data source that the template can reference.
        var data = new List<Item>
        {
            new Item { Id = 1, Name = "First",  Value = 10.0 },
            new Item { Id = 2, Name = "Second", Value = 20.0 },
            new Item { Id = 3, Name = "Third",  Value = 30.0 }
        };

        // Populate the template using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, data, "items"); // "items" is the name used in the template.

        // Configure image save options to render each page as a separate frame in a TIFF.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff);
        saveOptions.PageLayout = MultiPageLayout.TiffFrames(); // each page becomes a TIFF frame

        // Save the populated document as a multi‑frame TIFF file.
        doc.Save("Result.tiff", saveOptions);
    }

    // Simple POCO class used as the data source.
    public class Item
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public double Value { get; set; }
    }
}
