using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsLinqReporting
{
    // Simple data class that will be used as the LINQ data source.
    public class ReportItem
    {
        public string Name { get; set; }
        public int Value { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Load the PDF template that contains Aspose.Words reporting tags.
            // The template can reference the data source as <<foreach [items]>><<[Name]>><<[Value]>><</foreach>>.
            Document template = new Document("Template.pdf");

            // Prepare sequential data using LINQ (here a simple list of objects).
            List<ReportItem> data = new List<ReportItem>
            {
                new ReportItem { Name = "Alpha",   Value = 10 },
                new ReportItem { Name = "Beta",    Value = 20 },
                new ReportItem { Name = "Gamma",   Value = 30 },
                new ReportItem { Name = "Delta",   Value = 40 }
            };

            // Build the report by populating the template with the data source.
            // The data source name "items" must match the name used in the template tags.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(template, data, "items");

            // Configure image save options for PNG output.
            ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
            {
                // Optional: increase resolution for higher quality images.
                Resolution = 300
            };

            // Export each page of the populated document to a separate PNG file.
            for (int pageIndex = 0; pageIndex < template.PageCount; pageIndex++)
            {
                // Set the PageSet to render only the current page (zero‑based index).
                pngOptions.PageSet = new PageSet(pageIndex);

                // Define the output file name (e.g., Page_1.png, Page_2.png, ...).
                string outputPath = $"Page_{pageIndex + 1}.png";

                // Save the current page as a PNG image.
                template.Save(outputPath, pngOptions);
            }
        }
    }
}
