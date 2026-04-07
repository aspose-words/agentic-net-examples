using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;   // Required for Table type

public class Item
{
    public int Index { get; set; }
    public string Name { get; set; } = string.Empty;
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
        for (int i = 0; i < 10; i++)
        {
            model.Items.Add(new Item { Index = i, Name = $"Item {i}" });
        }

        // Create a template document with LINQ Reporting tags.
        string templatePath = Path.Combine(Environment.CurrentDirectory, "template.docx");
        CreateTemplate(templatePath);

        // Load the template.
        var doc = new Document(templatePath);

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string reportPath = Path.Combine(Environment.CurrentDirectory, "report.docx");
        doc.Save(reportPath);
    }

    private static void CreateTemplate(string filePath)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Begin the foreach loop over Items.
        builder.Writeln("<<foreach [item in Items]>>");

        // Create a table with two columns: Index and Name.
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
            "<<if [item.Index % 2 == 0]>><<backColor [\"LightGray\"]>><<[item.Index]>> <</backColor>><</if>>" +
            "<<if [item.Index % 2 != 0]>><<[item.Index]>> <</if>>");

        builder.InsertCell();
        builder.Writeln(
            "<<if [item.Index % 2 == 0]>><<backColor [\"LightGray\"]>><<[item.Name]>> <</backColor>><</if>>" +
            "<<if [item.Index % 2 != 0]>><<[item.Name]>> <</if>>");

        builder.EndRow();
        builder.EndTable();

        // End the foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(filePath);
    }
}
