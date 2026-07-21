using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables; // Needed for the Table class

namespace AsposeWordsLinqReportingBookmarks
{
    // Data model for the report.
    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public int Id { get; set; }
        public string Name { get; set; } = "";
        public string Description { get; set; } = "";
        // The bookmark name that will be used for this row.
        public string BookmarkName { get; set; } = "";
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            string templatePath = "Template.docx";
            string outputPath = "Report.docx";

            // -----------------------------------------------------------------
            // 1. Create the LINQ Reporting template programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Table header (static, not inside the foreach).
            builder.Writeln("<<foreach [item in Items]>>");
            Table table = builder.StartTable();

            // Header row.
            builder.InsertCell();
            builder.Writeln("ID");
            builder.InsertCell();
            builder.Writeln("Name");
            builder.InsertCell();
            builder.Writeln("Description");
            builder.EndRow();

            // Data row – each row will have a bookmark around the ID cell.
            builder.InsertCell();
            // Open bookmark tag, the expression returns the bookmark name for the current item.
            builder.Writeln("<<bookmark [item.BookmarkName]>>");
            // Write the ID inside the bookmark.
            builder.Writeln("<<[item.Id]>>");
            // Close bookmark tag.
            builder.Writeln("<</bookmark>>");

            builder.InsertCell();
            builder.Writeln("<<[item.Name]>>");
            builder.InsertCell();
            builder.Writeln("<<[item.Description]>>");
            builder.EndRow();

            builder.EndTable();
            builder.Writeln("<</foreach>>");

            // Add a list of hyperlinks that navigate to the bookmarks created above.
            builder.Writeln("Links to rows:");
            builder.Writeln("<<foreach [item in Items]>>");
            // The first expression is the bookmark target, the second is the display text.
            builder.Writeln("<<link [item.BookmarkName] [item.Name]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Prepare sample data.
            // -----------------------------------------------------------------
            var model = new ReportModel();

            for (int i = 1; i <= 5; i++)
            {
                model.Items.Add(new Item
                {
                    Id = i,
                    Name = $"Item {i}",
                    Description = $"Description for item {i}",
                    BookmarkName = $"bm_{i}"
                });
            }

            // -----------------------------------------------------------------
            // 3. Load the template and build the report using LINQ Reporting.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);
            ReportingEngine engine = new ReportingEngine();

            // Build the report. The root object name is "model" because the template uses <<[item ...]>> inside a foreach over Items.
            engine.BuildReport(reportDoc, model, "model");

            // -----------------------------------------------------------------
            // 4. Save the generated report.
            // -----------------------------------------------------------------
            reportDoc.Save(outputPath);
        }
    }
}
