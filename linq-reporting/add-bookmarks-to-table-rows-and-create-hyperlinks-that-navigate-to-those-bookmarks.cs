using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables; // Required for Table type

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Items = new List<RowItem>()
        };
        for (int i = 1; i <= 5; i++)
        {
            model.Items.Add(new RowItem
            {
                Id = i,
                Name = $"Item {i}",
                BookmarkName = $"bm_{i}"
            });
        }

        // -----------------------------------------------------------------
        // 1. Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Table header (outside the data band so it appears only once).
        builder.Writeln("<<foreach [item in Items]>>"); // Start data band.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Name (bookmarked)");
        builder.InsertCell();
        builder.Writeln("ID");
        builder.EndRow();

        // Table rows – each row gets a bookmark around the name cell.
        builder.InsertCell();
        builder.Writeln("<<bookmark [item.BookmarkName]>>");
        builder.Writeln("<<[item.Name]>>");
        builder.Writeln("<</bookmark>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Id]>>");
        builder.EndRow();

        builder.EndTable();
        builder.Writeln("<</foreach>>"); // End data band.

        builder.Writeln(); // Blank line.

        // List of hyperlinks that navigate to the bookmarks.
        builder.Writeln("Links to rows:");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("<<link [item.BookmarkName] [item.Name]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the final document.
        const string outputPath = "ReportWithBookmarks.docx";
        reportDoc.Save(outputPath);
    }
}

// ---------------------------------------------------------------------
// Data model used by the LINQ Reporting engine.
// ---------------------------------------------------------------------
public class ReportModel
{
    public List<RowItem> Items { get; set; } = new();
}

public class RowItem
{
    public int Id { get; set; }
    public string Name { get; set; } = "";
    public string BookmarkName { get; set; } = "";
}
