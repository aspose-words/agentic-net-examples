using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace BookmarkExtractionExample
{
    // Model representing a single bookmark entry.
    public class BookmarkItem
    {
        // Name of the bookmark (must be unique within the document).
        public string Name { get; set; } = string.Empty;

        // Text that will appear inside the bookmark.
        public string Title { get; set; } = string.Empty;
    }

    // Root model passed to the LINQ Reporting engine.
    public class ReportModel
    {
        // Collection of bookmark items to be rendered.
        public List<BookmarkItem> Bookmarks { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create a template document that contains LINQ Reporting tags.
            // -----------------------------------------------------------------
            var template = new Document();
            var builder = new DocumentBuilder(template);

            // Begin a foreach loop over the Bookmarks collection.
            builder.Writeln("<<foreach [b in Bookmarks]>>");

            // Define a bookmark whose name comes from the model and whose
            // content is the Title property.
            builder.Writeln("<<bookmark [b.Name]>>");
            builder.Writeln("<<[b.Title]>>");
            builder.Writeln("<</bookmark>>");

            // End the foreach loop.
            builder.Writeln("<</foreach>>");

            // Save the template to disk (required before BuildReport).
            const string templatePath = "Template.docx";
            template.Save(templatePath);

            // ---------------------------------------------------------------
            // 2. Prepare sample data that the report will be built from.
            // ---------------------------------------------------------------
            var model = new ReportModel
            {
                Bookmarks = new()
                {
                    new() { Name = "BM1", Title = "First Bookmark" },
                    new() { Name = "BM2", Title = "Second Bookmark" },
                    new() { Name = "BM3", Title = "Third Bookmark" }
                }
            };

            // ---------------------------------------------------------------
            // 3. Build the report using the ReportingEngine.
            // ---------------------------------------------------------------
            var reportDoc = new Document(templatePath);
            var engine = new ReportingEngine();

            // The root object name in the template is "model".
            engine.BuildReport(reportDoc, model, "model");

            // Save the generated report.
            const string reportPath = "Report.docx";
            reportDoc.Save(reportPath);

            // ---------------------------------------------------------------
            // 4. Extract bookmark names from the generated document.
            // ---------------------------------------------------------------
            var extractedNames = reportDoc.Range.Bookmarks
                                            .Select(b => b.Name)
                                            .ToList();

            // Expected bookmark names based on the model.
            var expectedNames = model.Bookmarks
                                     .Select(b => b.Name)
                                     .ToList();

            // ---------------------------------------------------------------
            // 5. Verify that the extracted names match the expected ones.
            // ---------------------------------------------------------------
            bool match = extractedNames.SequenceEqual(expectedNames);

            Console.WriteLine($"Bookmark extraction match: {match}");
            Console.WriteLine("Extracted bookmark names:");
            foreach (var name in extractedNames)
            {
                Console.WriteLine($"- {name}");
            }

            // The example finishes without waiting for user input.
        }
    }
}
