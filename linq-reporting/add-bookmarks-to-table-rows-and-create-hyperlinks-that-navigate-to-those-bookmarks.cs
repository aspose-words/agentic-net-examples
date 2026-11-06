using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create the template document.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Write a list of hyperlinks that point to the bookmarks.
        builder.Writeln("Links to rows:");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("<<link [item.BookmarkName] [item.Name]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln(); // Blank line.

        // Generate a table row for each item with a bookmark.
        builder.Writeln("<<foreach [item in Items]>>");
        Table table = builder.StartTable();

        // First cell contains the bookmark.
        builder.InsertCell();
        builder.Writeln("<<bookmark [item.BookmarkName]>>");
        builder.Writeln("<<[item.Name]>>");
        builder.Writeln("<</bookmark>>");

        // Second cell contains a description.
        builder.InsertCell();
        builder.Writeln("<<[item.Description]>>");

        builder.EndRow();
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Save the template.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // Prepare sample data.
        ReportModel model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Id = 1, Name = "First Item", Description = "Description of first item." },
                new Item { Id = 2, Name = "Second Item", Description = "Description of second item." },
                new Item { Id = 3, Name = "Third Item", Description = "Description of third item." }
            }
        };

        // Build the report.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };
        engine.BuildReport(reportDoc, model, "model");

        // Save the final document.
        reportDoc.Save("Report.docx");
    }
}

// Data model classes.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public int Id { get; set; }
    public string Name { get; set; } = "";
    public string Description { get; set; } = "";
    public string BookmarkName => $"bm_{Id}";
}
