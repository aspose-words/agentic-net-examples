using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Sample data model.
        var model = new ReportModel
        {
            Groups = new()
            {
                new Group
                {
                    Name = "Group A",
                    Items = new()
                    {
                        new Item { Description = "Item 1", Quantity = 10 },
                        new Item { Description = "Item 2", Quantity = 20 }
                    }
                },
                new Group
                {
                    Name = "Group B",
                    Items = new()
                    {
                        new Item { Description = "Item 3", Quantity = 30 }
                    }
                }
            }
        };

        // -----------------------------------------------------------------
        // Create the template document programmatically.
        // -----------------------------------------------------------------
        string templatePath = "Template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Outer foreach – iterate over groups.
        builder.Writeln("<<foreach [group in Groups]>>");

        // -------------------------------------------------------------
        // Header table – contains a single row with merged cells.
        // -------------------------------------------------------------
        Table headerTable = builder.StartTable();
        builder.InsertCell();
        builder.Write("<<cellMerge -horz>><<[group.Name]>>");
        builder.InsertCell();
        builder.Write("<<cellMerge -horz>><<[group.Name]>>");
        builder.EndRow();
        builder.EndTable();

        // -------------------------------------------------------------
        // Items tables – one table per item (simplified safe pattern).
        // -------------------------------------------------------------
        builder.Writeln("<<foreach [item in group.Items]>>");
        Table itemsTable = builder.StartTable();

        // Column titles (appears for each item – acceptable for demo).
        builder.InsertCell();
        builder.Writeln("Description");
        builder.InsertCell();
        builder.Writeln("Qty");
        builder.EndRow();

        // Data row.
        builder.InsertCell();
        builder.Writeln("<<[item.Description]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Quantity]>>");
        builder.EndRow();

        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Close outer foreach.
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Load the template and build the report.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report.
        string reportPath = "Report.docx";
        reportDoc.Save(reportPath);

        Console.WriteLine($"Report generated: {Path.GetFullPath(reportPath)}");
    }
}

// ---------------------------------------------------------------------
// Data model classes.
// ---------------------------------------------------------------------
public class ReportModel
{
    public List<Group> Groups { get; set; } = new();
}

public class Group
{
    public string Name { get; set; } = "";
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Description { get; set; } = "";
    public int Quantity { get; set; }
}
