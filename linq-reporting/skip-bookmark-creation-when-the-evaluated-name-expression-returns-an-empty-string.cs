using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create the template document with LINQ Reporting tags.
        var templatePath = "Template.docx";
        var outputPath = "Report.docx";

        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Begin foreach over Items collection.
        builder.Writeln("<<foreach [item in Items]>>");

        // If BookmarkName is not empty, create a bookmark around the title.
        builder.Writeln("<<if [item.BookmarkName != \"\"]>>");
        builder.Writeln("<<bookmark [item.BookmarkName]>>");
        builder.Writeln("<<[item.Title]>>");
        builder.Writeln("<</bookmark>>");
        builder.Writeln("<</if>>");

        // If BookmarkName is empty, just write the title without a bookmark.
        builder.Writeln("<<if [item.BookmarkName == \"\"]>>");
        builder.Writeln("<<[item.Title]>>");
        builder.Writeln("<</if>>");

        // End foreach.
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(templatePath);

        // Load the template for report generation.
        var template = new Document(templatePath);

        // Prepare sample data.
        var model = new ReportModel
        {
            Items = new()
            {
                new Item { Title = "First Item", BookmarkName = "FirstBookmark" },
                new Item { Title = "Second Item", BookmarkName = "" },
                new Item { Title = "Third Item", BookmarkName = "ThirdBookmark" }
            }
        };

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(template, model, "model");

        // Save the generated report.
        template.Save(outputPath);
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
    public string Title { get; set; } = "";
    public string BookmarkName { get; set; } = "";
}
