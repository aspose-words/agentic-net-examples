using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqBookmarkExample
{
    public class Program
    {
        public static void Main()
        {
            // Register code pages provider (required for some Aspose.Words features).
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            // -------------------------------------------------
            // 1. Create a template document with LINQ Reporting tags.
            // -------------------------------------------------
            var template = new Document();
            var builder = new DocumentBuilder(template);

            // Define a foreach loop that creates a bookmark for each item.
            builder.Writeln("<<foreach [item in Items]>>");
            builder.Writeln("<<bookmark [item.BookmarkName]>>");
            builder.Writeln("<<[item.Title]>>");
            builder.Writeln("<</bookmark>>");
            builder.Writeln("<</foreach>>");

            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // -------------------------------------------------
            // 2. Load the template and build the report.
            // -------------------------------------------------
            var doc = new Document(templatePath);

            var model = new ReportModel
            {
                Items = new List<Item>
                {
                    new Item { Id = 1, Title = "First Item" },
                    new Item { Id = 2, Title = "Second Item" },
                    new Item { Id = 3, Title = "Third Item" }
                }
            };

            var engine = new ReportingEngine();
            // The root object name in the template is "model".
            engine.BuildReport(doc, model, "model");

            const string reportPath = "Report.docx";
            doc.Save(reportPath);

            // -------------------------------------------------
            // 3. Output created bookmark names (optional verification).
            // -------------------------------------------------
            foreach (Bookmark bookmark in doc.Range.Bookmarks)
            {
                Console.WriteLine($"Bookmark: {bookmark.Name}, Text: {bookmark.Text.Trim()}");
            }
        }
    }

    // Root data model for the report.
    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    // Individual item that provides a dynamic bookmark name.
    public class Item
    {
        public int Id { get; set; }
        public string Title { get; set; } = "";
        // Bookmark name derived from the Id field.
        public string BookmarkName => $"Item_{Id}";
    }
}
