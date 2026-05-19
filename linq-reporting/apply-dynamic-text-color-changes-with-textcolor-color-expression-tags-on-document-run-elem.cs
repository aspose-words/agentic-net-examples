using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Data model for the report.
    public class ReportModel
    {
        // Collection of items to be displayed.
        public List<Item> Items { get; set; } = new();
    }

    // Individual item containing the text and the color name.
    public class Item
    {
        public string Text { get; set; } = "";
        public string ColorName { get; set; } = "";
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            var template = new Document();
            var builder = new DocumentBuilder(template);

            // Insert a foreach block that iterates over Items.
            builder.Writeln("<<foreach [item in Items]>>");
            // Apply a dynamic text color based on the item's ColorName property.
            builder.Writeln("<<textColor [item.ColorName]>><<[item.Text]>><</textColor>>");
            builder.Writeln("<</foreach>>");

            // Save the template to a local file.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Load the template back (required before building the report).
            // -----------------------------------------------------------------
            var doc = new Document(templatePath);

            // -----------------------------------------------------------------
            // 3. Prepare the data source.
            // -----------------------------------------------------------------
            var model = new ReportModel
            {
                Items = new List<Item>
                {
                    new Item { Text = "Operation succeeded", ColorName = "Green" },
                    new Item { Text = "Warning: Check input", ColorName = "Orange" },
                    new Item { Text = "Error occurred", ColorName = "Red" }
                }
            };

            // -----------------------------------------------------------------
            // 4. Build the report using the LINQ Reporting engine.
            // -----------------------------------------------------------------
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            const string outputPath = "ReportWithDynamicColors.docx";
            doc.Save(outputPath);
        }
    }
}
