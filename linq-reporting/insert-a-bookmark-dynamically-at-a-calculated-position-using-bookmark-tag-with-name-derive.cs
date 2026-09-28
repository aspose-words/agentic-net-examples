using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some environments)
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        string templatePath = "template.docx";
        string outputPath = "report.docx";

        // Create the LINQ Reporting template
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Begin foreach loop over Items collection
        builder.Writeln("<<foreach [item in Items]>>");
        // Insert a bookmark with a dynamic name derived from data
        builder.Writeln("<<bookmark [item.BookmarkName]>>");
        // Content that will be inside the bookmark
        builder.Writeln("<<[item.Title]>>");
        // Close bookmark tag
        builder.Writeln("<</bookmark>>");
        // End foreach loop
        builder.Writeln("<</foreach>>");

        // Save the template to disk
        templateDoc.Save(templatePath);

        // Load the template for report generation
        var reportDoc = new Document(templatePath);

        // Prepare sample data
        var model = new ReportModel
        {
            Items = new()
            {
                new Item { Title = "Introduction", BookmarkName = "bm_Intro" },
                new Item { Title = "Chapter 1", BookmarkName = "bm_Chapter1" },
                new Item { Title = "Conclusion", BookmarkName = "bm_Conclusion" }
            }
        };

        // Build the report using LINQ Reporting engine
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report
        reportDoc.Save(outputPath);
    }
}

// Root data model
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

// Item model used in the foreach loop
public class Item
{
    public string Title { get; set; } = "";
    public string BookmarkName { get; set; } = "";
}
