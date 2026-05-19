using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;   // Required for Table type

public class Program
{
    public static void Main()
    {
        // Sample data.
        var model = new ReportModel
        {
            Items = new()
            {
                new Item { Group = "Group A", Name = "Item 1" },
                new Item { Group = "Group A", Name = "Item 2" },
                new Item { Group = "Group B", Name = "Item 3" }
            }
        };

        // Create a template document programmatically.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Begin foreach block.
        builder.Writeln("<<foreach [item in Items]>>");

        // Build a table for each item.
        Table table = builder.StartTable();

        // First row – merged cells showing the group name.
        builder.InsertCell();
        builder.Write("<<cellMerge>><<[item.Group]>>");
        builder.InsertCell();
        builder.Write("<<cellMerge>><<[item.Group]>>");
        builder.EndRow();

        // Second row – regular cells with item details.
        builder.InsertCell();
        builder.Write("<<[item.Name]>>");
        builder.InsertCell();
        builder.Write("Details");
        builder.EndRow();

        builder.EndTable();

        // End foreach block.
        builder.Writeln("<</foreach>>");

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the result.
        doc.Save("MergedTableReport.docx");
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
    public string Group { get; set; } = "";
    public string Name { get; set; } = "";
}
