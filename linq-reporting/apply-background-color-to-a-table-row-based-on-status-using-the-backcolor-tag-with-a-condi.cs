using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some environments)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create the data model
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "Task A", Status = "Completed" },
                new Item { Name = "Task B", Status = "InProgress" },
                new Item { Name = "Task C", Status = "Pending" }
            }
        };

        // Build the template document with LINQ Reporting tags
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Task Report");
        builder.Writeln("<<foreach [item in Items]>>");

        // Table header
        var table = builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Status");
        builder.EndRow();

        // Table row with conditional background color
        builder.InsertCell();
        builder.Writeln(
            "<<if [item.Status == \"Completed\"]>><<backColor [\"LightGreen\"]>><<[item.Name]>> <</backColor>><</if>>" +
            "<<if [item.Status != \"Completed\"]>><<[item.Name]>> <</if>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Status]>>");
        builder.EndRow();

        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Generate the report
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the result
        const string outputPath = "ReportWithConditionalBackColor.docx";
        doc.Save(outputPath);
        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}

// Data model classes
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = string.Empty;
    public string Status { get; set; } = string.Empty;
}
