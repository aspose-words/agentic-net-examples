using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Required for some Aspose.Words features.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Sample data model.
        ReportModel model = new()
        {
            Groups = new List<Group>
            {
                new()
                {
                    Name = "Fruits",
                    Items = new List<Item>
                    {
                        new() { Name = "Apple", Quantity = 10 },
                        new() { Name = "Banana", Quantity = 20 }
                    }
                },
                new()
                {
                    Name = "Vegetables",
                    Items = new List<Item>
                    {
                        new() { Name = "Carrot", Quantity = 15 },
                        new() { Name = "Tomato", Quantity = 25 }
                    }
                }
            }
        };

        // -----------------------------------------------------------------
        // Create the template document programmatically.
        // -----------------------------------------------------------------
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        Document template = new();
        DocumentBuilder builder = new(template);

        // Outer foreach – iterate over groups.
        builder.Writeln("<<foreach [group in Groups]>>");

        // Table with a merged row that shows the group name.
        Table groupTable = builder.StartTable();
        builder.InsertCell();
        builder.Write("<<cellMerge -horz>><<[group.Name]>>");
        builder.InsertCell();
        builder.Write("<<cellMerge -horz>><<[group.Name]>>");
        builder.EndRow();
        builder.EndTable();

        // Header row for the items table.
        Table headerTable = builder.StartTable();
        builder.InsertCell();
        builder.Write("Item");
        builder.InsertCell();
        builder.Write("Quantity");
        builder.EndRow();
        builder.EndTable();

        // Inner foreach – iterate over items of the current group.
        builder.Writeln("<<foreach [item in group.Items]>>");

        // Each item is rendered as a one‑row table (keeps the table structure safe).
        Table itemTable = builder.StartTable();
        builder.InsertCell();
        builder.Write("<<[item.Name]>>");
        builder.InsertCell();
        builder.Write("<<[item.Quantity]>>");
        builder.EndRow();
        builder.EndTable();

        builder.Writeln("<</foreach>>"); // close inner foreach
        builder.Writeln("<</foreach>>"); // close outer foreach

        // Save the template.
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // Build the report using the template and the data model.
        // -----------------------------------------------------------------
        Document report = new(templatePath);
        ReportingEngine engine = new()
        {
            Options = ReportBuildOptions.None
        };
        engine.BuildReport(report, model, "model");

        // Save the final report.
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        report.Save(reportPath);
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
    public string Name { get; set; } = string.Empty;
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = string.Empty;
    public int Quantity { get; set; }
}
