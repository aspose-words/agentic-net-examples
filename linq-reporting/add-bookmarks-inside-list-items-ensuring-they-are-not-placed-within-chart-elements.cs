using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingBookmarks
{
    // Data model for the report.
    public class ReportModel
    {
        // Collection of items to be listed.
        public List<Item> Items { get; set; } = new();
    }

    // Individual list item.
    public class Item
    {
        // Text displayed for the item.
        public string Title { get; set; } = string.Empty;

        // Name of the bookmark that will be placed around the item.
        public string BookmarkName { get; set; } = string.Empty;
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

            // Apply a numbered list style to subsequent paragraphs.
            builder.ListFormat.List = template.Lists.Add(ListTemplate.NumberDefault);
            builder.Writeln("Report Items:");

            // LINQ Reporting tags:
            //   <<foreach [item in Items]>> ... <</foreach>>
            //   <<bookmark [item.BookmarkName]>> ... <</bookmark>>
            builder.Writeln("<<foreach [item in Items]>>");
            builder.Writeln("<<bookmark [item.BookmarkName]>>");
            builder.Writeln("<<[item.Title]>>");
            builder.Writeln("<</bookmark>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Prepare sample data.
            // -----------------------------------------------------------------
            var model = new ReportModel
            {
                Items = new List<Item>
                {
                    new Item { Title = "First list item", BookmarkName = "bmFirst" },
                    new Item { Title = "Second list item", BookmarkName = "bmSecond" },
                    new Item { Title = "Third list item", BookmarkName = "bmThird" }
                }
            };

            // -----------------------------------------------------------------
            // 3. Load the template and build the report.
            // -----------------------------------------------------------------
            var doc = new Document(templatePath);
            var engine = new ReportingEngine();

            // Build the report using the model as the root object named "model".
            engine.BuildReport(doc, model, "model");

            // Save the generated report.
            const string outputPath = "Report.docx";
            doc.Save(outputPath);
        }
    }
}
