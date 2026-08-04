using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

namespace AsposeWordsLinqReportingBookmarks
{
    // Data model for each row.
    public class Item
    {
        public int Id { get; set; }
        public string Name { get; set; } = string.Empty;

        // Bookmark name derived from the Id.
        public string BookmarkName => $"Row_{Id}";
    }

    // Root object passed to the ReportingEngine.
    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            const string templatePath = "Template.docx";
            const string reportPath = "Report.docx";

            // -------------------------------------------------
            // 1. Create the template document programmatically.
            // -------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Report with bookmarks and hyperlinks:");

            // Begin the foreach block. The whole table will be generated for each item.
            builder.Writeln("<<foreach [item in Items]>>");

            // Start a table for the current item.
            Table table = builder.StartTable();

            // First cell: bookmark that surrounds the item name.
            builder.InsertCell();
            builder.Writeln("<<bookmark [item.BookmarkName]>> <<[item.Name]>> <</bookmark>>");

            // Second cell: hyperlink that points to the same bookmark.
            builder.InsertCell();
            builder.Writeln("<<link [item.BookmarkName] [\"Go to \" + item.Name]>>");

            // End the row for the current item.
            builder.EndRow();

            // Close the table for the current item.
            builder.EndTable();

            // End the foreach block.
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -------------------------------------------------
            // 2. Load the template and prepare the data source.
            // -------------------------------------------------
            Document doc = new Document(templatePath);

            ReportModel model = new ReportModel
            {
                Items = new List<Item>
                {
                    new Item { Id = 1, Name = "Alpha" },
                    new Item { Id = 2, Name = "Beta" },
                    new Item { Id = 3, Name = "Gamma" }
                }
            };

            // -------------------------------------------------
            // 3. Build the report using the LINQ Reporting engine.
            // -------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // -------------------------------------------------
            // 4. Save the generated report.
            // -------------------------------------------------
            doc.Save(reportPath);
        }
    }
}
