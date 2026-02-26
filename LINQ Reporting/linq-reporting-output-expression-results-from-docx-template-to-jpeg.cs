using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsLinqReporting
{
    // Simple data model used as the data source for the LINQ Reporting Engine.
    public class Item
    {
        public int Value { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Load the DOCX template that contains LINQ expressions, e.g.
            // <<[data.Items.Where(i => i.Value > 10).Count()]>>
            Document doc = new Document("Template.docx");

            // Prepare the data source.
            var dataSource = new
            {
                Items = new List<Item>
                {
                    new Item { Value = 5 },
                    new Item { Value = 12 },
                    new Item { Value = 20 }
                }
            };

            // Build the report using the LINQ Reporting Engine.
            ReportingEngine engine = new ReportingEngine();
            // The third parameter ("data") is the name used to reference the data source inside the template.
            engine.BuildReport(doc, dataSource, "data");

            // Configure image save options for JPEG output.
            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Jpeg)
            {
                // Render the first page of the document (zero‑based index).
                PageSet = new PageSet(0),

                // Optional: set JPEG quality (0‑100). Higher values give better quality.
                JpegQuality = 90
            };

            // Save the populated document as a JPEG image.
            doc.Save("Report.jpg", saveOptions);
        }
    }
}
