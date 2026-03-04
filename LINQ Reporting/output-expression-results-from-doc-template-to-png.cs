using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the template document that contains Aspose.Words reporting tags.
        string templatePath = @"C:\Docs\Template.docx";

        // Path where the rendered PNG image will be saved.
        string outputImagePath = @"C:\Docs\Result.png";

        // Load the template document.
        Document template = new Document(templatePath);

        // Example data source – replace with your actual data.
        var dataSource = new
        {
            Title = "Quarterly Report",
            Date = DateTime.Now,
            TotalSales = 123456.78
        };

        // Populate the template with the data source using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, dataSource);

        // Configure image save options to render the first page as a PNG.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Png)
        {
            // Render only the first page (zero‑based index).
            PageSet = new PageSet(0),

            // Optional: set resolution (dpi) if higher quality is required.
            Resolution = 300
        };

        // Save the populated document as a PNG image.
        template.Save(outputImagePath, saveOptions);
    }
}
