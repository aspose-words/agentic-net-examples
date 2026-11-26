using System;
using System.Collections.Generic;
using System.IO;
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
        // Prepare sample data.
        var model = new ReportModel();
        for (int i = 1; i <= 10; i++)
        {
            model.Items.Add(new Item { Index = i, Name = $"Item {i}" });
        }

        // Create the LINQ Reporting template.
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Begin foreach over Items.
        builder.Writeln("<<foreach [item in Items]>>");

        // Create a table with header.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Index");
        builder.InsertCell();
        builder.Writeln("Name");
        builder.EndRow();

        // Data row with conditional background color for even rows.
        builder.InsertCell();
        builder.Writeln(
            "<<if [item.Index % 2 == 0]>><<backColor [\"LightGray\"]>><<[item.Index]>> <</backColor>><</if>>" +
            "<<if [item.Index % 2 != 0]>><<[item.Index]>><</if>>");

        builder.InsertCell();
        builder.Writeln(
            "<<if [item.Index % 2 == 0]>><<backColor [\"LightGray\"]>><<[item.Name]>> <</backColor>><</if>>" +
            "<<if [item.Index % 2 != 0]>><<[item.Name]>><</if>>");

        builder.EndRow();
        builder.EndTable();

        // End foreach.
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(templatePath);

        // Load the template for report generation.
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();

        // Build the report.
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report.
        reportDoc.Save("Report.docx");
    }
}
