using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

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
            Items = new()
            {
                new Item { Name = "Task A", Status = "Completed" },
                new Item { Name = "Task B", Status = "Pending" },
                new Item { Name = "Task C", Status = "Completed" },
                new Item { Name = "Task D", Status = "InProgress" }
            }
        };

        // Create the template document.
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Title.
        builder.Writeln("Task Status Report");
        builder.Writeln();

        // Begin foreach loop over Items.
        builder.Writeln("<<foreach [item in Items]>>");

        // Table for each item (single row).
        Table table = builder.StartTable();

        // Name cell with conditional background color.
        builder.InsertCell();
        builder.Writeln(
            "<<if [item.Status == \"Completed\"]>><<backColor [\"LightGreen\"]>><<[item.Name]>> <</backColor>><</if>>" +
            "<<if [item.Status != \"Completed\"]>><<[item.Name]>> <</if>>");

        // Status cell with conditional background color.
        builder.InsertCell();
        builder.Writeln(
            "<<if [item.Status == \"Completed\"]>><<backColor [\"LightGreen\"]>><<[item.Status]>> <</backColor>><</if>>" +
            "<<if [item.Status != \"Completed\"]>><<[item.Status]>> <</if>>");

        builder.EndRow();
        builder.EndTable();

        // End foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(templatePath);

        // Load the template for reporting.
        var templateDoc = new Document(templatePath);

        // Build the report.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(templateDoc, model, "model");

        // Save the generated report.
        var outputPath = "ReportOutput.docx";
        templateDoc.Save(outputPath);
    }
}
