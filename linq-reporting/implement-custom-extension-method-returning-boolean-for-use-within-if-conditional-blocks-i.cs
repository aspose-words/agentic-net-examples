using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = "";

    // Custom method usable in LINQ Reporting templates
    public bool IsEven() => Index % 2 == 0;
}

public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Sample data
        var model = new ReportModel();
        for (int i = 1; i <= 5; i++)
        {
            model.Items.Add(new Item { Index = i, Name = $"Item {i}" });
        }

        // Create template
        const string templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Begin foreach block
        builder.Writeln("<<foreach [item in Items]>>");

        // Table header
        Table table = builder.StartTable();
        builder.InsertCell(); builder.Writeln("Index");
        builder.InsertCell(); builder.Writeln("Name");
        builder.EndRow();

        // Data row
        builder.InsertCell();
        builder.Writeln(
            "<<if [item.IsEven()]>>" +
            "<<backColor [\"LightGray\"]>><<[item.Index]>> <</backColor>><</if>>" +
            "<<if [item.IsEven() == false]>>" +
            "<<[item.Index]>> <</if>>");
        builder.InsertCell();
        builder.Writeln(
            "<<if [item.IsEven()]>>" +
            "<<backColor [\"LightGray\"]>><<[item.Name]>> <</backColor>><</if>>" +
            "<<if [item.IsEven() == false]>>" +
            "<<[item.Name]>> <</if>>");
        builder.EndRow();

        // End table and foreach
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Save the template
        doc.Save(templatePath);

        // Build report
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");
        reportDoc.Save("Report.docx");
    }
}
