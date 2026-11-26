using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Group
{
    public string Name { get; set; } = "";
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public int Index { get; set; }
    public string Description { get; set; } = "";
}

public class ReportModel
{
    public List<Group> Groups { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create the template document with LINQ Reporting tags.
        // -----------------------------------------------------------------
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Outer foreach – iterate over groups.
        builder.Writeln("<<foreach [group in Groups]>>");

        // Row that displays the group name merged across two cells.
        Table groupTable = builder.StartTable();
        builder.InsertCell();
        builder.Write("<<cellMerge -horz>><<[group.Name]>>");
        builder.InsertCell();
        builder.Write("<<cellMerge -horz>><<[group.Name]>>");
        builder.EndRow();
        builder.EndTable();

        // Inner foreach – iterate over items of the current group.
        builder.Writeln("<<foreach [item in group.Items]>>");

        // Table that holds the items.
        Table itemTable = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Index");
        builder.InsertCell();
        builder.Writeln("Description");
        builder.EndRow();

        // Data row – repeated for each item.
        builder.InsertCell();
        builder.Writeln("<<[item.Index]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Description]>>");
        builder.EndRow();

        builder.EndTable();
        builder.Writeln("<</foreach>>"); // Close inner foreach.

        // Close outer foreach.
        builder.Writeln("<</foreach>>");

        // Save the template.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);

        // Sample data.
        var model = new ReportModel
        {
            Groups = new()
            {
                new Group
                {
                    Name = "Group A",
                    Items = new()
                    {
                        new Item { Index = 1, Description = "Item A1" },
                        new Item { Index = 2, Description = "Item A2" }
                    }
                },
                new Group
                {
                    Name = "Group B",
                    Items = new()
                    {
                        new Item { Index = 1, Description = "Item B1" },
                        new Item { Index = 2, Description = "Item B2" },
                        new Item { Index = 3, Description = "Item B3" }
                    }
                }
            }
        };

        // Build the report.
        var engine = new ReportingEngine
        {
            Options = ReportBuildOptions.None
        };
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report.
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        reportDoc.Save(reportPath);
    }
}
