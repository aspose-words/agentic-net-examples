using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace BookmarkExtractionExample
{
    // Data model for the LINQ Reporting template.
    public class ReportModel
    {
        // Collection of items that will be iterated over in the template.
        public List<Item> Items { get; set; } = new();
    }

    // Individual item containing a bookmark name and some text.
    public class Item
    {
        public string BookmarkName { get; set; } = "";
        public string Title { get; set; } = "";
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Prepare sample data.
            // -----------------------------------------------------------------
            var model = new ReportModel
            {
                Items =
                {
                    new Item { BookmarkName = "FirstBookmark", Title = "First Title" },
                    new Item { BookmarkName = "SecondBookmark", Title = "Second Title" },
                    new Item { BookmarkName = "ThirdBookmark", Title = "Third Title" }
                }
            };

            // -----------------------------------------------------------------
            // 2. Create a template document that uses LINQ Reporting tags.
            // -----------------------------------------------------------------
            var template = new Document();
            var builder = new DocumentBuilder(template);

            // Begin a foreach loop over the Items collection.
            builder.Writeln("<<foreach [item in Items]>>");
            // Insert a bookmark whose name comes from the current item.
            builder.Writeln("<<bookmark [item.BookmarkName]>>");
            // The content of the bookmark – the title text.
            builder.Writeln("<<[item.Title]>>");
            // Close the bookmark.
            builder.Writeln("<</bookmark>>");
            // End the foreach loop.
            builder.Writeln("<</foreach>>");

            // -----------------------------------------------------------------
            // 3. Build the report using the ReportingEngine.
            // -----------------------------------------------------------------
            var engine = new ReportingEngine();
            engine.BuildReport(template, model, "model");

            // -----------------------------------------------------------------
            // 4. Extract bookmark names from the generated document.
            // -----------------------------------------------------------------
            List<string> extractedNames = template.Range.Bookmarks
                                                    .Select(b => b.Name)
                                                    .ToList();

            // -----------------------------------------------------------------
            // 5. Verify that the extracted names match the expected ones.
            // -----------------------------------------------------------------
            List<string> expectedNames = model.Items
                                             .Select(i => i.BookmarkName)
                                             .ToList();

            bool match = extractedNames.SequenceEqual(expectedNames);

            // Output the verification result.
            Console.WriteLine("Extracted bookmark names:");
            foreach (var name in extractedNames)
                Console.WriteLine($"- {name}");

            Console.WriteLine();
            Console.WriteLine($"Verification {(match ? "succeeded" : "failed")}.");
        }
    }
}
