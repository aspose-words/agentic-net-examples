using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables; // Needed for Table type

public class Item
{
    public string Name { get; set; } = "";
    public string Status { get; set; } = "";
}

public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Sample data.
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "Task A", Status = "Completed" },
                new Item { Name = "Task B", Status = "Pending" },
                new Item { Name = "Task C", Status = "Completed" },
                new Item { Name = "Task D", Status = "InProgress" }
            }
        };

        // Build the template document.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Begin foreach loop.
        builder.Writeln("<<foreach [item in Items]>>");

        // Table definition.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Status");
        builder.EndRow();

        // Data row with conditional background color.
        builder.InsertCell();
        builder.Writeln(
            "<<if [item.Status == \"Completed\"]>><<backColor [\"LightGreen\"]>><<[item.Name]>> <</backColor>><</if>>" +
            "<<if [item.Status != \"Completed\"]>><<[item.Name]>> <</if>>");
        builder.InsertCell();
        builder.Writeln(
            "<<if [item.Status == \"Completed\"]>><<backColor [\"LightGreen\"]>><<[item.Status]>> <</backColor>><</if>>" +
            "<<if [item.Status != \"Completed\"]>><<[item.Status]>> <</if>>");
        builder.EndRow();

        // End table and foreach.
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Generate the report.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save output.
        doc.Save("Report.docx");
    }
}
