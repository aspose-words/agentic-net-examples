using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class BookmarkItem
{
    public string BookmarkName { get; set; } = "";
    public string InnerBookmarkName { get; set; } = "";
    public string Title { get; set; } = "";
}

public class ReportModel
{
    public List<BookmarkItem> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare template document with nested tables and bookmark tags.
        var templatePath = "Template.docx";
        var reportPath = "Report.docx";

        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Begin foreach over Items collection.
        builder.Writeln("<<foreach [item in Items]>>");

        // Outer table.
        Table outerTable = builder.StartTable();
        builder.InsertCell();

        // Outer bookmark.
        builder.Writeln("<<bookmark [item.BookmarkName]>>");
        builder.Writeln("<<[item.Title]>>");
        builder.Writeln("<</bookmark>>");

        // Inner table inside the same outer cell.
        Table innerTable = builder.StartTable();
        builder.InsertCell();

        // Inner bookmark.
        builder.Writeln("<<bookmark [item.InnerBookmarkName]>>");
        builder.Writeln("Inner <<[item.Title]>>");
        builder.Writeln("<</bookmark>>");

        builder.EndRow();          // End inner table row.
        builder.EndTable();        // End inner table.

        builder.EndRow();          // End outer table row.
        builder.EndTable();        // End outer table.

        // End foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Create sample data.
        var model = new ReportModel
        {
            Items = new List<BookmarkItem>
            {
                new()
                {
                    BookmarkName = "OuterBookmark1",
                    InnerBookmarkName = "InnerBookmark1",
                    Title = "First Item"
                },
                new()
                {
                    BookmarkName = "OuterBookmark2",
                    InnerBookmarkName = "InnerBookmark2",
                    Title = "Second Item"
                }
            }
        };

        // Load the template for reporting.
        var reportDoc = new Document(templatePath);

        // Build the report.
        var engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report.
        reportDoc.Save(reportPath);
    }
}
