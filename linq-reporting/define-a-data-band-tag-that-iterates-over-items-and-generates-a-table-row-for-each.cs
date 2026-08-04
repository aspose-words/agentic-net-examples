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
        ReportModel model = new()
        {
            Items = new()
            {
                new Item { Index = 1, Name = "Apple" },
                new Item { Index = 2, Name = "Banana" },
                new Item { Index = 3, Name = "Cherry" }
            }
        };

        // Build the template document.
        Document template = new();
        DocumentBuilder builder = new(template);

        // Title.
        builder.Writeln("Items Report");
        builder.Writeln();

        // Data band: iterate over Items and generate a row for each.
        builder.Writeln("<<foreach [item in Items]>>");

        // Table with header and data rows.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Index");
        builder.InsertCell();
        builder.Writeln("Name");
        builder.EndRow();

        // Data row (repeated for each item).
        builder.InsertCell();
        builder.Writeln("<<[item.Index]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Name]>>");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Close the foreach block.
        builder.Writeln("<</foreach>>");

        // Build the report.
        ReportingEngine engine = new();
        engine.BuildReport(template, model, "model");

        // Save the generated document.
        template.Save("Report.docx");
    }
}

// Data model classes.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = string.Empty;
}
