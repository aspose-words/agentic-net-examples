using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a blank document that will serve as the template.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Template: iterate over Items and create a bookmark only when BookmarkName is not empty.
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("<<if [!string.IsNullOrEmpty(item.BookmarkName)]>>");
        builder.Writeln("<<bookmark [item.BookmarkName]>>");
        builder.Writeln("<<[item.Title]>>");
        builder.Writeln("<</bookmark>>");
        builder.Writeln("<</if>>");
        builder.Writeln("<</foreach>>");

        // Prepare sample data.
        ReportModel model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Title = "First Item", BookmarkName = "BM1" },
                new Item { Title = "Second Item", BookmarkName = "" } // Empty name – bookmark will be skipped.
            }
        };

        // Build the report.
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.RemoveEmptyParagraphs
        };
        engine.BuildReport(doc, model, "model");

        // Save the result.
        doc.Save("Report.docx");
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
    public string Title { get; set; } = string.Empty;
    public string BookmarkName { get; set; } = string.Empty;
}
