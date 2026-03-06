using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the DOCX template that contains Aspose.Words reporting tags.
        string templatePath = @"C:\Templates\ReportTemplate.docx";

        // Load the template document.
        Document doc = new Document(templatePath);

        // Prepare a data source for the template.
        // This can be any POCO, DataSet, etc. Here we use an anonymous object for simplicity.
        var dataSource = new
        {
            Title = "Quarterly Sales",
            Date  = DateTime.Now,
            Total = 123456.78
        };

        // Populate the template with data using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, dataSource);

        // Configure image save options for JPEG output.
        ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
        {
            // Optional: set JPEG quality (0‑100). 100 = best quality, larger file size.
            JpegQuality = 90,

            // Optional: enable high‑quality rendering for better visual fidelity.
            UseHighQualityRendering = true,

            // Optional: render only the first page (default behavior for image formats).
            // If you need a specific page, set the PageSet property, e.g.:
            // PageSet = new PageSet(0) // zero‑based index of the page to render.
        };

        // Path for the resulting JPEG image.
        string outputPath = @"C:\Output\ReportPage1.jpg";

        // Save the populated document as a JPEG image.
        doc.Save(outputPath, jpegOptions);

        Console.WriteLine("Report rendered to JPEG successfully.");
    }
}
