using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Items = new List<Item>
            {
                new Item { Name = "Item 1", Status = "A" },
                new Item { Name = "Item 2", Status = "B" },
                new Item { Name = "Item 3", Status = "C" } // Will trigger the default case.
            }
        };

        // Create a template document programmatically.
        var template = new Document();
        var builder = new DocumentBuilder(template);

        builder.Writeln("Report:");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("- <<[item.Name]>>: ");

        // Conditional blocks for known statuses.
        builder.Writeln("<<if [item.Status == \"A\"]>>Alpha<</if>>");
        builder.Writeln("<<if [item.Status == \"B\"]>>Beta<</if>>");

        // Default block when none of the above conditions are true.
        builder.Writeln("<<if [item.Status != \"A\" && item.Status != \"B\"]>>Other<</if>>");

        builder.Writeln("<</foreach>>");

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(template, model, "model");

        // Save the generated report.
        template.Save("Report.docx");
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
