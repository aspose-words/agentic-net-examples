using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the PDF template that contains LINQ Reporting tags.
        string templatePath = "Template.pdf";

        // Directory where the generated JPEG images will be saved.
        string outputDir = "OutputImages";
        Directory.CreateDirectory(outputDir);

        // Load the PDF template into an Aspose.Words Document.
        Document doc = new Document(templatePath);

        // Prepare a data source that the template can iterate over.
        // The template should reference this collection with the name "items".
        var data = new List<ReportItem>
        {
            new ReportItem { Id = 1, Name = "First", Value = 123.45 },
            new ReportItem { Id = 2, Name = "Second", Value = 678.90 },
            new ReportItem { Id = 3, Name = "Third", Value = 1011.12 }
        };

        // Populate the template using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, data, "items");

        // Render each page of the populated document to a separate JPEG file.
        for (int pageIndex = 0; pageIndex < doc.PageCount; pageIndex++)
        {
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Jpeg)
            {
                // Render only the current page.
                PageSet = new PageSet(pageIndex),

                // Adjust JPEG quality as needed (0‑100).
                JpegQuality = 90
            };

            string outFile = Path.Combine(outputDir, $"Page_{pageIndex + 1}.jpg");
            doc.Save(outFile, options);
        }
    }

    // Simple POCO class used as a data source for the report.
    public class ReportItem
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public double Value { get; set; }
    }
}
