using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Data model classes
    public class ReportModel
    {
        // Collection of items to be iterated in the template
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public string Name { get; set; } = string.Empty;
        // Description may be null or empty; such paragraphs will be removed
        public string? Description { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the temporary template and final report
            string templatePath = Path.Combine(Environment.CurrentDirectory, "Template.docx");
            string reportPath   = Path.Combine(Environment.CurrentDirectory, "Report.docx");

            // -------------------------------------------------
            // 1. Create the template document programmatically
            // -------------------------------------------------
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Add a title
            builder.Writeln("Report of Items");
            builder.Writeln();

            // LINQ Reporting tags:
            // Iterate over Items collection; each iteration writes Name and Description on separate paragraphs.
            builder.Writeln("<<foreach [item in Items]>>");
            builder.Writeln("Name: <<[item.Name]>>");
            // Description may be empty; after processing the paragraph could become empty.
            builder.Writeln("Description: <<[item.Description]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk (required before loading for reporting)
            template.Save(templatePath);

            // -------------------------------------------------
            // 2. Load the template for report generation
            // -------------------------------------------------
            Document doc = new Document(templatePath);

            // -------------------------------------------------
            // 3. Prepare sample data
            // -------------------------------------------------
            ReportModel model = new ReportModel
            {
                Items = new List<Item>
                {
                    new Item { Name = "Item A", Description = "First item description." },
                    new Item { Name = "Item B", Description = null },               // Will produce an empty paragraph
                    new Item { Name = "Item C", Description = "" },                 // Will produce an empty paragraph
                    new Item { Name = "Item D", Description = "Fourth item description." }
                }
            };

            // -------------------------------------------------
            // 4. Build the report with removal of empty paragraphs
            // -------------------------------------------------
            ReportingEngine engine = new ReportingEngine
            {
                // Enable removal of paragraphs that become empty after tag processing
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };

            // Build the report; the root object name must match the tag prefix ("model")
            engine.BuildReport(doc, model, "model");

            // -------------------------------------------------
            // 5. Save the final document
            // -------------------------------------------------
            doc.Save(reportPath);

            // Optional: indicate completion (no interactive input)
            Console.WriteLine($"Report generated successfully: {reportPath}");
        }
    }
}
