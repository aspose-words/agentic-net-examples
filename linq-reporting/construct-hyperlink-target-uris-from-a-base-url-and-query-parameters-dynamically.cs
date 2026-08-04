using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace HyperlinkLinqReportingExample
{
    // Data model for the report.
    public class ReportModel
    {
        // Base URL used to build hyperlink targets.
        public string BaseUrl { get; set; } = string.Empty;

        // Collection of items to iterate over.
        public List<Item> Items { get; set; } = new();
    }

    // Individual item containing parameters for the hyperlink.
    public class Item
    {
        public int Id { get; set; }
        public string Name { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create the template document programmatically.
            var template = new Document();
            var builder = new DocumentBuilder(template);

            // Begin a foreach loop over the Items collection.
            builder.Writeln("<<foreach [item in model.Items]>>");

            // Insert a link tag where the URI is built from BaseUrl and the item's Id,
            // and the display text is the item's Name.
            builder.Writeln("<<link [model.BaseUrl + \"?id=\" + item.Id] [item.Name]>>");

            // End the foreach loop.
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // 2. Load the template for reporting.
            var document = new Document(templatePath);

            // 3. Prepare sample data.
            var model = new ReportModel
            {
                BaseUrl = "https://example.com/product",
                Items = new List<Item>
                {
                    new Item { Id = 101, Name = "Widget A" },
                    new Item { Id = 202, Name = "Gadget B" },
                    new Item { Id = 303, Name = "Doohickey C" }
                }
            };

            // 4. Build the report using the LINQ Reporting engine.
            var engine = new ReportingEngine();
            engine.BuildReport(document, model, "model");

            // 5. Save the generated report.
            const string outputPath = "ReportWithHyperlinks.docx";
            document.Save(outputPath);
        }
    }
}
