using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace BookmarkInListExample
{
    // Data model for the report.
    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public string Name { get; set; } = "";
        public string BookmarkName { get; set; } = "";
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider for Aspose.Words.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Paths for template and output.
            string templatePath = "Template.docx";
            string outputPath = "Report.docx";

            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Write LINQ Reporting tags for a list with bookmarks.
            builder.Writeln("<<foreach [item in Items]>>");
            // Each list item will have a bookmark whose name comes from the data source.
            builder.Writeln("<<bookmark [item.BookmarkName]>>");
            // The actual list content (item name) goes between the bookmark tags.
            builder.Writeln("<<[item.Name]>>");
            builder.Writeln("<</bookmark>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Prepare sample data.
            // -----------------------------------------------------------------
            ReportModel model = new ReportModel
            {
                Items = new List<Item>
                {
                    new Item { Name = "First item", BookmarkName = "bmFirst" },
                    new Item { Name = "Second item", BookmarkName = "bmSecond" },
                    new Item { Name = "Third item", BookmarkName = "bmThird" }
                }
            };

            // -----------------------------------------------------------------
            // 3. Load the template and build the report.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(reportDoc, model, "model");

            // Save the generated report.
            reportDoc.Save(outputPath);
        }
    }
}
