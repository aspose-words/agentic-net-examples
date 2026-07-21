using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingBookmarks
{
    // Simple data model for the report.
    public class ReportModel
    {
        // Initialize the collection to avoid nullable warnings.
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public string BookmarkName { get; set; } = string.Empty;
        public string Title { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a template document that uses LINQ Reporting tags to
            //    generate bookmarks.
            // -----------------------------------------------------------------
            var templatePath = "Template.docx";
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            // Begin a foreach loop over the Items collection.
            builder.Writeln("<<foreach [item in Items]>>");
            // Open a bookmark whose name comes from the data source.
            builder.Writeln("<<bookmark [item.BookmarkName]>>");
            // Insert the title inside the bookmark.
            builder.Writeln("<<[item.Title]>>");
            // Close the bookmark.
            builder.Writeln("<</bookmark>>");
            // End the foreach loop.
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Prepare sample data that will be merged into the template.
            // -----------------------------------------------------------------
            var model = new ReportModel();
            model.Items.Add(new Item { BookmarkName = "FirstBookmark", Title = "First Item" });
            model.Items.Add(new Item { BookmarkName = "SecondBookmark", Title = "Second Item" });
            model.Items.Add(new Item { BookmarkName = "ThirdBookmark", Title = "Third Item" });

            // -----------------------------------------------------------------
            // 3. Load the template and build the report using the LINQ Reporting engine.
            // -----------------------------------------------------------------
            var reportDoc = new Document(templatePath);
            var engine = new ReportingEngine();
            // No special options are required for this scenario.
            engine.BuildReport(reportDoc, model, "model");

            // -----------------------------------------------------------------
            // 4. Extract bookmark names from the generated document.
            // -----------------------------------------------------------------
            List<string> actualBookmarkNames = reportDoc.Range.Bookmarks
                .Select(b => b.Name)
                .ToList();

            // Expected bookmark names come from the source data.
            List<string> expectedBookmarkNames = model.Items
                .Select(i => i.BookmarkName)
                .ToList();

            // -----------------------------------------------------------------
            // 5. Verify that the extracted bookmark names match the expected ones.
            // -----------------------------------------------------------------
            bool match = actualBookmarkNames.SequenceEqual(expectedBookmarkNames);

            Console.WriteLine("Expected bookmark names:");
            foreach (var name in expectedBookmarkNames)
                Console.WriteLine($"  {name}");

            Console.WriteLine("\nActual bookmark names:");
            foreach (var name in actualBookmarkNames)
                Console.WriteLine($"  {name}");

            Console.WriteLine($"\nVerification result: {(match ? "SUCCESS" : "FAILURE")}");

            // Save the final document for inspection (optional).
            var outputPath = "ReportWithBookmarks.docx";
            reportDoc.Save(outputPath);
        }
    }
}
