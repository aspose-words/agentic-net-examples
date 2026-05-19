using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace HyperlinkValidationExample
{
    // Data model for the report.
    public class ReportModel
    {
        // Collection of items that will be iterated in the template.
        public List<Item> Items { get; set; } = new();
    }

    // Individual item containing hyperlink data.
    public class Item
    {
        // Target of the hyperlink. Must not be empty.
        public string Url { get; set; } = string.Empty;

        // Text displayed for the hyperlink.
        public string Text { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            string templatePath = "Template.docx";
            string outputPath = "Report.docx";

            // -----------------------------------------------------------------
            // 1. Create the LINQ Reporting template programmatically.
            // -----------------------------------------------------------------
            DocumentBuilder builder = new DocumentBuilder();
            builder.Writeln("Hyperlink Report");
            builder.Writeln("<<foreach [item in Items]>>");
            // Link tag: first expression is the target, second is the display text.
            builder.Writeln("<<link [item.Url] [item.Text]>>");
            builder.Writeln("<</foreach>>");
            builder.Document.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Prepare sample data, intentionally leaving some URLs empty.
            // -----------------------------------------------------------------
            ReportModel model = new()
            {
                Items = new List<Item>
                {
                    new Item { Url = "https://www.example.com", Text = "Example Site" },
                    new Item { Url = "", Text = "Missing URL" },               // Invalid entry
                    new Item { Url = "https://www.github.com", Text = "GitHub" },
                    new Item { Url = null, Text = "Null URL" }                // Invalid entry
                }
            };

            // -----------------------------------------------------------------
            // 3. Validate hyperlink targets before building the report.
            // -----------------------------------------------------------------
            bool hasErrors = false;
            foreach (var item in model.Items)
            {
                if (string.IsNullOrWhiteSpace(item.Url))
                {
                    Console.WriteLine($"Error: Hyperlink target is empty for display text \"{item.Text}\".");
                    hasErrors = true;
                }
            }

            if (hasErrors)
            {
                Console.WriteLine("Report generation aborted due to hyperlink validation errors.");
                return;
            }

            // -----------------------------------------------------------------
            // 4. Load the template and build the report.
            // -----------------------------------------------------------------
            Document template = new(templatePath);
            ReportingEngine engine = new ReportingEngine();
            // No special options are required for this scenario.
            engine.BuildReport(template, model, "model");

            // -----------------------------------------------------------------
            // 5. Save the generated report.
            // -----------------------------------------------------------------
            template.Save(outputPath);
            Console.WriteLine($"Report generated successfully: {Path.GetFullPath(outputPath)}");
        }
    }
}
