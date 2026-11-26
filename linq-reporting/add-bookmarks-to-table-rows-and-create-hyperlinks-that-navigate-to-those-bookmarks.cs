using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

namespace AsposeWordsLinqReportingBookmarks
{
    // Data model for a table row.
    public class RowItem
    {
        public int Id { get; set; }
        public string Name { get; set; } = "";
        // Bookmark name must be unique and non‑empty.
        public string BookmarkName => $"Row{Id}";
    }

    // Root model passed to the reporting engine.
    public class ReportModel
    {
        public List<RowItem> Items { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some environments).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare sample data.
            var model = new ReportModel
            {
                Items = new List<RowItem>
                {
                    new RowItem { Id = 1, Name = "Alpha" },
                    new RowItem { Id = 2, Name = "Beta" },
                    new RowItem { Id = 3, Name = "Gamma" }
                }
            };

            // -----------------------------------------------------------------
            // Create the LINQ Reporting template programmatically.
            // -----------------------------------------------------------------
            var templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Section with hyperlinks that navigate to the bookmarks.
            builder.Writeln("Links to rows:");
            builder.Writeln("<<foreach [item in Items]>>");
            builder.Writeln("<<link [item.BookmarkName] [item.Name]>>");
            builder.Writeln("<</foreach>>");
            builder.Writeln(); // Blank line.

            // Table rows, each row is a separate table containing a bookmark.
            builder.Writeln("<<foreach [item in Items]>>");
            Table table = builder.StartTable();

            // First cell – contains the bookmark that wraps the Id.
            builder.InsertCell();
            builder.Writeln("<<bookmark [item.BookmarkName]>>");
            builder.Writeln("<<[item.Id]>>");
            builder.Writeln("<</bookmark>>");

            // Second cell – displays the Name.
            builder.InsertCell();
            builder.Writeln("<<[item.Name]>>");

            builder.EndRow();
            builder.EndTable();
            builder.Writeln("<</foreach>>");

            // Save the template.
            doc.Save(templatePath);

            // -----------------------------------------------------------------
            // Load the template and build the report.
            // -----------------------------------------------------------------
            var reportDoc = new Document(templatePath);
            var engine = new ReportingEngine();
            engine.BuildReport(reportDoc, model, "model");

            // Save the final report.
            var outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
            reportDoc.Save(outputPath);
        }
    }
}
