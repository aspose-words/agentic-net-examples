using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsMarkdownExport
{
    // Sample data class used as a data source for the template.
    public class ReportData
    {
        public string Title { get; set; }
        public List<Item> Items { get; set; }

        public ReportData()
        {
            Items = new List<Item>();
        }
    }

    public class Item
    {
        public string Name { get; set; }
        public int Quantity { get; set; }
    }

    public static class MarkdownExporter
    {
        // Exports a DOCX template populated with data to a Markdown file.
        public static void ExportToMarkdown(string templatePath, ReportData data, string outputPath)
        {
            // Load the DOCX template.
            Document doc = new Document(templatePath);

            // Populate the template using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, data, "data");

            // Configure Markdown save options.
            MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
            {
                SaveFormat = SaveFormat.Markdown,
                // Optional: export tables as raw HTML if needed.
                // ExportAsHtml = MarkdownExportAsHtml.Tables,
                // Optional: export OfficeMath as LaTeX.
                // OfficeMathExportMode = MarkdownOfficeMathExportMode.Latex
            };

            // Save the populated document as Markdown.
            doc.Save(outputPath, saveOptions);
        }

        // Example usage.
        public static void Main()
        {
            // Path to the DOCX template containing Aspose.Words tags, e.g. <<[data.Title]>> and <<foreach [data.Items]>><<[Name]>> - <<[Quantity]>>><</foreach>>.
            string templatePath = @"C:\Templates\ReportTemplate.docx";

            // Prepare sample data.
            ReportData data = new ReportData
            {
                Title = "Monthly Inventory Report",
                Items = new List<Item>
                {
                    new Item { Name = "Apples", Quantity = 120 },
                    new Item { Name = "Bananas", Quantity = 85 },
                    new Item { Name = "Oranges", Quantity = 60 }
                }
            };

            // Destination Markdown file.
            string outputPath = @"C:\Output\Report.md";

            // Perform the export.
            ExportToMarkdown(templatePath, data, outputPath);

            Console.WriteLine("Export completed: " + outputPath);
        }
    }
}
