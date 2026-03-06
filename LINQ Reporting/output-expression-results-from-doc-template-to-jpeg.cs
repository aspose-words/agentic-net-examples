using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the DOC template that contains expression tags (e.g. <<[Data.Name]>>)
        Document template = new Document("Template.docx");

        // Prepare the data source that will be merged into the template.
        // This can be any POCO, DataSet, etc. Here we use an anonymous object for simplicity.
        var dataSource = new
        {
            Name = "John Doe",
            Date = DateTime.Now,
            Amount = 1234.56
        };

        // Populate the template with the data using ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        bool success = engine.BuildReport(template, dataSource, "Data");
        if (!success)
        {
            Console.WriteLine("Template processing failed.");
            return;
        }

        // Configure image save options for JPEG output.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Jpeg)
        {
            // Adjust quality if needed (0‑100). 95 is default.
            JpegQuality = 90,
            // Optional: set resolution (dpi) for higher quality rendering.
            Resolution = 300,
            // Optional: render only the first page (default renders first page only for images).
            // PageSet = new PageSet(0);
        };

        // Save the populated document as a JPEG image.
        template.Save("Result.jpg", saveOptions);

        Console.WriteLine("Document rendered to JPEG successfully.");
    }
}
