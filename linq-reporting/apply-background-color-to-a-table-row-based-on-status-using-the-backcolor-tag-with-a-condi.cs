using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create the template document.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Begin a foreach loop over the Items collection.
        builder.Writeln("<<foreach [item in Items]>>");

        // Create a table with a header row.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Status");
        builder.EndRow();

        // Data row – apply background color based on the item's Status.
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

        // End the foreach block.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        string templatePath = "Template.docx";
        template.Save(templatePath);

        // Prepare sample data.
        ReportModel model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "Task A", Status = "Completed" },
                new Item { Name = "Task B", Status = "Pending" },
                new Item { Name = "Task C", Status = "Completed" },
                new Item { Name = "Task D", Status = "InProgress" }
            }
        };

        // Load the template and build the report.
        Document report = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };
        engine.BuildReport(report, model, "model");

        // Save the generated report.
        report.Save("Report.docx");
    }
}

// Root data model.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

// Item model used in the foreach loop.
public class Item
{
    public string Name { get; set; } = string.Empty;
    public string Status { get; set; } = string.Empty;
}
