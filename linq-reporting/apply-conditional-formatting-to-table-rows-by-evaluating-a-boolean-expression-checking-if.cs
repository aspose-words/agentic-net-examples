using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

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
        // ---------- Create the template document ----------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Title
        builder.Writeln("Items Report");

        // Begin the foreach loop that will repeat the table rows.
        builder.Writeln("<<foreach [item in Items]>>");

        // Create the table that will be repeated for each item.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Index");
        builder.InsertCell();
        builder.Writeln("Name");
        builder.EndRow();

        // Data row with conditional background color for even rows.
        builder.InsertCell();
        builder.Writeln(
            "<<if [item.Index % 2 == 0]>>" +
            "<<backColor [\"LightGray\"]>><<[item.Index]>><</backColor>><</if>>" +
            "<<if [item.Index % 2 != 0]>>" +
            "<<[item.Index]>>" +
            "<</if>>");

        builder.InsertCell();
        builder.Writeln(
            "<<if [item.Index % 2 == 0]>>" +
            "<<backColor [\"LightGray\"]>><<[item.Name]>><</backColor>><</if>>" +
            "<<if [item.Index % 2 != 0]>>" +
            "<<[item.Name]>>" +
            "<</if>>");

        // End the data row.
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // End the foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template.
        const string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // ---------- Build the report ----------
        var reportDoc = new Document(templatePath);

        var model = new ReportModel
        {
            Items = new()
            {
                new Item { Index = 0, Name = "Alice" },
                new Item { Index = 1, Name = "Bob" },
                new Item { Index = 2, Name = "Charlie" },
                new Item { Index = 3, Name = "Diana" },
                new Item { Index = 4, Name = "Ethan" }
            }
        };

        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        const string reportPath = "Report.docx";
        reportDoc.Save(reportPath);
    }
}
