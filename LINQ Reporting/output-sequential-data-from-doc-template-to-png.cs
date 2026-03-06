using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsExample
{
    class Program
    {
        static void Main()
        {
            // Path to the Word template that contains reporting tags (e.g. <<[Data.Name]>>)
            string templatePath = @"C:\Docs\Template.docx";

            // Load the template document (lifecycle rule: load)
            Document doc = new Document(templatePath);

            // Prepare a simple data source that matches the tags in the template.
            // The object can be any POCO; here we use an anonymous type for brevity.
            var dataSource = new
            {
                Name = "John Doe",
                Address = "123 Main St, Anytown",
                Items = new[]
                {
                    new { Description = "Item A", Quantity = 2, Price = 9.99 },
                    new { Description = "Item B", Quantity = 1, Price = 19.95 }
                }
            };

            // Populate the template with data using the ReportingEngine (feature rule).
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, dataSource, "Data");

            // Render each page of the populated document to a separate PNG file.
            // Use ImageSaveOptions with SaveFormat.Png (feature rule for image saving).
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png);

            // Optional: set resolution or image size if required.
            // options.Resolution = 300; // DPI
            // options.ImageSize = new System.Drawing.Size(1240, 1754); // A4 at 300 DPI

            // Iterate through all pages and save them sequentially.
            for (int pageIndex = 0; pageIndex < doc.PageCount; pageIndex++)
            {
                // Set the page to render (zero‑based index) – this follows the example rule.
                options.PageSet = new PageSet(pageIndex);

                // Build the output file name, e.g. Output_1.png, Output_2.png, …
                string outputPath = Path.Combine(
                    @"C:\Docs\OutputImages",
                    $"Output_{pageIndex + 1}.png");

                // Save the current page as PNG (lifecycle rule: save)
                doc.Save(outputPath, options);
            }
        }
    }
}
