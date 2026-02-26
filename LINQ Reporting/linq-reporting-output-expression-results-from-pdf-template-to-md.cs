using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsLinqReportingExample
{
    // Simple data source class used in the LINQ reporting template.
    public class ReportData
    {
        public string Title { get; set; }
        public List<Item> Items { get; set; }

        public class Item
        {
            public string Name { get; set; }
            public decimal Price { get; set; }
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Path to the input PDF template that contains LINQ reporting tags.
            const string templatePath = @"InputTemplate.pdf";

            // Path to the output Markdown file.
            const string outputPath = @"ReportOutput.md";

            // Load the PDF template into an Aspose.Words Document.
            Document doc = new Document(templatePath);

            // Prepare the data source for the report.
            var data = new ReportData
            {
                Title = "Product Catalog",
                Items = new List<ReportData.Item>
                {
                    new ReportData.Item { Name = "Apple",  Price = 0.99m },
                    new ReportData.Item { Name = "Banana", Price = 0.59m },
                    new ReportData.Item { Name = "Cherry", Price = 2.49m }
                }
            };

            // Create and configure the ReportingEngine.
            var engine = new ReportingEngine
            {
                // Allow missing members so the engine does not throw if a tag is not matched.
                Options = ReportBuildOptions.AllowMissingMembers
            };

            // Build the report by merging the data source with the template.
            // The second overload allows referencing the data source object itself via the name "ds".
            engine.BuildReport(doc, data, "ds");

            // Configure Markdown save options.
            var mdOptions = new MarkdownSaveOptions
            {
                // Export any OfficeMath as plain text (default) – can be changed if needed.
                OfficeMathExportMode = MarkdownOfficeMathExportMode.Text,
                // Preserve page breaks to keep the logical sections from the template.
                ForcePageBreaks = true,
                // Use UTF‑8 encoding.
                Encoding = System.Text.Encoding.UTF8
            };

            // Save the populated document as Markdown.
            doc.Save(outputPath, mdOptions);
        }
    }
}
