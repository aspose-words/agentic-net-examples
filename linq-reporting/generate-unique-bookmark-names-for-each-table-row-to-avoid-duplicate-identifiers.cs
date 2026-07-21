using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

namespace AsposeWordsLinqReportingDemo
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
        // Unique bookmark name for each row.
        public string BookmarkName { get; set; } = "";
    }

    public class Program
    {
        public static void Main()
        {
            // Paths for the template and the generated report.
            string templatePath = Path.Combine(Environment.CurrentDirectory, "Template.docx");
            string reportPath = Path.Combine(Environment.CurrentDirectory, "Report.docx");

            // -----------------------------------------------------------------
            // 1. Create the template document programmatically.
            // -----------------------------------------------------------------
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            // Begin a foreach loop over Items.
            builder.Writeln("<<foreach [item in Items]>>");

            // Create a table that will be repeated for each item.
            Table table = builder.StartTable();

            // Header row.
            builder.InsertCell();
            builder.Writeln("ID");
            builder.InsertCell();
            builder.Writeln("Name");
            builder.EndRow();

            // Data row – each cell contains LINQ Reporting tags.
            // The first cell wraps the ID inside a unique bookmark.
            builder.InsertCell();
            builder.Writeln("<<bookmark [item.BookmarkName]>>");
            builder.Writeln("<<[item.Id]>>");
            builder.Writeln("<</bookmark>>");

            // The second cell shows the item name.
            builder.InsertCell();
            builder.Writeln("<<[item.Name]>>");
            builder.EndRow();

            // End the table before closing the foreach block.
            builder.EndTable();

            // Close the foreach loop.
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            templateDoc.Save(templatePath);

            // -----------------------------------------------------------------
            // 2. Prepare sample data with unique bookmark names.
            // -----------------------------------------------------------------
            ReportModel model = new ReportModel();

            for (int i = 1; i <= 5; i++)
            {
                model.Items.Add(new Item
                {
                    Id = i,
                    Name = $"Item {i}",
                    // Generate a unique bookmark name per row.
                    BookmarkName = $"RowBookmark_{i}"
                });
            }

            // -----------------------------------------------------------------
            // 3. Load the template and build the report.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templatePath);

            ReportingEngine engine = new ReportingEngine();
            // BuildReport requires the root object name to match the tags (model).
            engine.BuildReport(reportDoc, model, "model");

            // Save the generated report.
            reportDoc.Save(reportPath);
        }
    }
}
