using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Id = 1, Name = "Alpha" },
                new Item { Id = 2, Name = "Beta" },
                new Item { Id = 3, Name = "Gamma" }
            }
        };

        // Create the LINQ Reporting template programmatically.
        var templatePath = "Template.docx";
        CreateTemplate(templatePath);

        // Load the template.
        var doc = new Document(templatePath);

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated document.
        var outputPath = "Report.docx";
        doc.Save(outputPath);
    }

    private static void CreateTemplate(string filePath)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Title.
        builder.Writeln("Items Table:");
        builder.Writeln();

        // Begin foreach loop over Items.
        builder.Writeln("<<foreach [item in Items]>>");

        // Start table.
        var table = builder.StartTable();

        // First cell – contains the bookmark start, Id and bookmark end.
        builder.InsertCell();
        builder.Writeln("<<bookmark [item.BookmarkName]>><<[item.Id]>><</bookmark>>");

        // Second cell – Name.
        builder.InsertCell();
        builder.Writeln("<<[item.Name]>>");

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // End foreach loop.
        builder.Writeln("<</foreach>>");
        builder.Writeln();

        // Links section.
        builder.Writeln("Links to rows:");
        builder.Writeln();

        // Begin another foreach to create hyperlinks.
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("<<link [item.BookmarkName] [item.LinkText]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(filePath);
    }
}

// Data model for the report.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

// Individual item displayed in the table.
public class Item
{
    public int Id { get; set; }
    public string Name { get; set; } = string.Empty;

    // Bookmark name generated from the Id.
    public string BookmarkName => $"Row_{Id}";

    // Text displayed for the hyperlink.
    public string LinkText => $"Go to {Name}";
}
