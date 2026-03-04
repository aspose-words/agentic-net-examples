using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsExample
{
    class Program
    {
        static void Main()
        {
            // Path to the DOCX template that contains Aspose.Words reporting tags.
            const string templatePath = @"C:\Templates\ReportTemplate.docx";

            // Load the template document (lifecycle rule: use Document constructor).
            Document doc = new Document(templatePath);

            // Example data source – replace with your own model.
            var data = new
            {
                Title = "Quarterly Sales",
                Date = DateTime.Now,
                Items = new[]
                {
                    new { Product = "Laptop", Quantity = 12, Price = 899.99 },
                    new { Product = "Smartphone", Quantity = 30, Price = 499.50 },
                    new { Product = "Tablet", Quantity = 20, Price = 299.00 }
                }
            };

            // Populate the template using the ReportingEngine (feature rule: BuildReport).
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, data, "ds");

            // Render each page of the populated document to a separate PNG file.
            // Use ImageSaveOptions (Save(*string, SaveOptions) overload) to specify PNG format and page range.
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
            {
                // Optional: set resolution or image size if needed.
                Resolution = 300
            };

            for (int i = 0; i < doc.PageCount; i++)
            {
                // Render only the current page.
                options.PageSet = new PageSet(i);

                // Save the page as PNG (lifecycle rule: use Document.Save).
                string outputPath = $@"C:\Output\Report_Page_{i + 1}.png";
                doc.Save(outputPath, options);
            }
        }
    }
}
