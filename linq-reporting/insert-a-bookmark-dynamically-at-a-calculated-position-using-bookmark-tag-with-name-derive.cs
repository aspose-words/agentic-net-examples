using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare folders.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Paths for template and result.
        string templatePath = Path.Combine(outputDir, "template.docx");
        string resultPath = Path.Combine(outputDir, "result.docx");

        // ---------- Create template ----------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Dynamic Bookmark Report");
        builder.Writeln();

        // Begin foreach over Items collection.
        builder.Writeln("<<foreach [item in Items]>>");

        // Insert a bookmark whose name is derived from the data field.
        builder.Writeln("<<bookmark [item.BookmarkName]>>");
        // Content that will be placed inside the bookmark.
        builder.Writeln("<<[item.Content]>>");
        // Close bookmark.
        builder.Writeln("<</bookmark>>");

        // End foreach.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // ---------- Prepare data ----------
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "Alpha", Content = "Content for Alpha" },
                new Item { Name = "Beta", Content = "Content for Beta" },
                new Item { Name = "Gamma", Content = "Content for Gamma" }
            }
        };

        // ---------- Build report ----------
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated document.
        reportDoc.Save(resultPath);
    }
}

// Root data model.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

// Item model used in the foreach loop.
public class Item
{
    public string Name { get; set; } = "";
    public string Content { get; set; } = "";

    // Bookmark name derived from the Name field.
    public string BookmarkName => "bm_" + Name;
}
