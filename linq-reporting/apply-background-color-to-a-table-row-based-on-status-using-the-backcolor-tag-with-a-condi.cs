using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

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

        // Create template document.
        const string templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Begin foreach loop for data rows.
        builder.Writeln("<<foreach [item in Items]>>");

        // Create a table for each item (header + data row).
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
        builder.EndTable();

        // End foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(templatePath);

        // Load the template for reporting.
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();

        // Build the report.
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        const string outputPath = "ReportOutput.docx";
        reportDoc.Save(outputPath);
    }
}

// Data model classes.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = "";
    public string Status { get; set; } = "";
}
