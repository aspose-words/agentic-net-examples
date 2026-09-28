using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Sample data model.
        var model = new ReportModel
        {
            Items = new()
            {
                new Item { Name = "Item 1", BookmarkName = "bm1" },
                new Item { Name = "Item 2", BookmarkName = "bm2" },
                new Item { Name = "Item 3", BookmarkName = "bm3" }
            }
        };

        // Paths for the template and final report.
        const string templatePath = "Template.docx";
        const string outputPath = "Report.docx";

        // -----------------------------------------------------------------
        // Create the LINQ Reporting template.
        // -----------------------------------------------------------------
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Hyperlink list that points to the bookmarks.
        builder.Writeln("Links:");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("<<link [item.BookmarkName] [item.Name]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln(string.Empty);

        // Header table (single header row).
        builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Item");
        builder.EndRow();
        builder.EndTable();

        // Data rows – each row is a separate table containing a bookmark.
        builder.Writeln("<<foreach [item in Items]>>");
        builder.StartTable();
        builder.InsertCell();
        builder.Writeln("<<bookmark [item.BookmarkName]>>");
        builder.Writeln("<<[item.Name]>>");
        builder.Writeln("<</bookmark>>");
        builder.EndRow();
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        doc.Save(templatePath);

        // -----------------------------------------------------------------
        // Build the report using the LINQ Reporting engine.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report.
        reportDoc.Save(outputPath);
    }
}

// ---------------------------------------------------------------------
// Data model classes.
// ---------------------------------------------------------------------
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = string.Empty;
    public string BookmarkName { get; set; } = string.Empty;
}
