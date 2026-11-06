using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

namespace LinqReportingBookmarkExample
{
    // Data model for the report.
    public class ReportModel
    {
        public List<RowData> Rows { get; set; } = new();
    }

    public class RowData
    {
        public int Index { get; set; }
        public string Name { get; set; } = "";
        // Unique bookmark name for this row.
        public string BookmarkName { get; set; } = "";
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare sample data.
            var model = new ReportModel();
            for (int i = 1; i <= 5; i++)
            {
                model.Rows.Add(new RowData
                {
                    Index = i,
                    Name = $"Item {i}",
                    BookmarkName = $"Row_{i}"
                });
            }

            // Create the LINQ Reporting template programmatically.
            var templatePath = "Template.docx";
            CreateTemplate(templatePath);

            // Load the template.
            var doc = new Document(templatePath);

            // Build the report.
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model, "Model");

            // Save the generated report.
            var outputPath = "Report.docx";
            doc.Save(outputPath);
        }

        private static void CreateTemplate(string filePath)
        {
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Title.
            builder.Writeln("Table of Rows with Unique Bookmarks:");
            builder.Writeln();

            // Begin foreach loop over rows.
            builder.Writeln("<<foreach [row in Model.Rows]>>");

            // Start table inside the loop.
            Table table = builder.StartTable();

            // Header row (added only once, before the first data row).
            builder.InsertCell();
            builder.Writeln("Index");
            builder.InsertCell();
            builder.Writeln("Name");
            builder.EndRow();

            // Data row.
            builder.InsertCell();
            // Bookmark around the index value.
            builder.Writeln("<<bookmark [row.BookmarkName]>> <<[row.Index]>> <</bookmark>>");
            builder.InsertCell();
            builder.Writeln("<<[row.Name]>>");
            builder.EndRow();

            // End table.
            builder.EndTable();

            // End foreach loop.
            builder.Writeln("<</foreach>>");

            // Save the template.
            doc.Save(filePath);
        }
    }
}
