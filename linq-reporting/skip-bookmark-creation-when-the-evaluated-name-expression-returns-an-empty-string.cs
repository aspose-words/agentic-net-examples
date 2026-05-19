using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some environments)
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            // Prepare the data model with some items having empty bookmark names
            var model = new ReportModel
            {
                Items = new List<Item>
                {
                    new Item { Title = "First Item", BookmarkName = "FirstBookmark" },
                    new Item { Title = "Second Item", BookmarkName = "" }, // No bookmark should be created
                    new Item { Title = "Third Item", BookmarkName = "ThirdBookmark" }
                }
            };

            // Create the template document programmatically
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // LINQ Reporting tags:
            // - Iterate over Items
            // - Conditionally create a bookmark only when BookmarkName is not empty
            builder.Writeln("<<foreach [item in Items]>>");
            builder.Writeln("<<if [item.BookmarkName != \"\"]>>");
            builder.Writeln("<<bookmark [item.BookmarkName]>>");
            builder.Writeln("<<[item.Title]>>");
            builder.Writeln("<</bookmark>>");
            builder.Writeln("<</if>>");
            builder.Writeln("<</foreach>>");

            // Build the report
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;
            engine.BuildReport(template, model, "model");

            // Save the generated document
            template.Save("ReportWithConditionalBookmarks.docx");
        }
    }

    // Root data model
    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    // Item model used in the foreach loop
    public class Item
    {
        public string Title { get; set; } = "";
        public string BookmarkName { get; set; } = "";
    }
}
