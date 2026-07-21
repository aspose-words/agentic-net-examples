using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

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
        // Prepare sample data.
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "Task A", Status = "Completed" },
                new Item { Name = "Task B", Status = "Pending" },
                new Item { Name = "Task C", Status = "Failed" }
            }
        };

        // -----------------------------------------------------------------
        // Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Begin foreach loop over Items.
        builder.Writeln("<<foreach [item in Items]>>");

        // Create a table with two columns: Name and Status.
        var table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Status");
        builder.EndRow();

        // Data row.
        builder.InsertCell();
        builder.Writeln("<<[item.Name]>>");
        builder.InsertCell();

        // Apply background color based on item.Status using backColor tag and conditional expressions.
        builder.Writeln(
            "<<if [item.Status == \"Completed\"]>><<backColor [\"LightGreen\"]>><<[item.Status]>> <</backColor>><</if>>" +
            "<<if [item.Status == \"Pending\"]>><<backColor [\"LightYellow\"]>><<[item.Status]>> <</backColor>><</if>>" +
            "<<if [item.Status != \"Completed\" && item.Status != \"Pending\"]>><<backColor [\"LightCoral\"]>><<[item.Status]>> <</backColor>><</if>>");

        builder.EndRow();
        builder.EndTable();

        // End foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template to a file.
        const string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template and build the report.
        // -----------------------------------------------------------------
        var loadedDoc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(loadedDoc, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        loadedDoc.Save(outputPath);
    }
}
