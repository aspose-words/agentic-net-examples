using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables; // Required for Table type

public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = "";
}

public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Create a blank document that will serve as the template.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Begin a foreach block that iterates over the Items collection.
        builder.Writeln("<<foreach [item in Items]>>");

        // Create a table with a header row.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Index");
        builder.InsertCell();
        builder.Writeln("Name");
        builder.EndRow();

        // Data row – apply a yellow background to rows where Index is even.
        builder.InsertCell();
        builder.Writeln(
            "<<if [item.Index % 2 == 0]>><<backColor [\"Yellow\"]>><<[item.Index]>> <</backColor>><</if>>" +
            "<<if [item.Index % 2 != 0]>><<[item.Index]>><</if>>");
        builder.InsertCell();
        builder.Writeln(
            "<<if [item.Index % 2 == 0]>><<backColor [\"Yellow\"]>><<[item.Name]>> <</backColor>><</if>>" +
            "<<if [item.Index % 2 != 0]>><<[item.Name]>><</if>>");
        builder.EndRow();

        // Close the table and the foreach block.
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Prepare sample data.
        ReportModel model = new ReportModel();
        model.Items.Add(new Item { Index = 1, Name = "Alpha" });
        model.Items.Add(new Item { Index = 2, Name = "Beta" });
        model.Items.Add(new Item { Index = 3, Name = "Gamma" });
        model.Items.Add(new Item { Index = 4, Name = "Delta" });

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, model, "model");

        // Save the generated document.
        template.Save("ReportWithDynamicBackColor.docx");
    }
}
