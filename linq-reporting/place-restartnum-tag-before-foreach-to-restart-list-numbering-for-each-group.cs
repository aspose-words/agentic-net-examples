using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Groups = new List<Group>
            {
                new Group
                {
                    Name = "Group A",
                    Items = new List<Item>
                    {
                        new Item { Name = "Item A1" },
                        new Item { Name = "Item A2" },
                        new Item { Name = "Item A3" }
                    }
                },
                new Group
                {
                    Name = "Group B",
                    Items = new List<Item>
                    {
                        new Item { Name = "Item B1" },
                        new Item { Name = "Item B2" }
                    }
                }
            }
        };

        // Create the template document programmatically.
        var templatePath = "Template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Outer foreach for groups.
        builder.Writeln("<<foreach [group in Groups]>>");
        // Group heading.
        builder.Writeln("Group: <<[group.Name]>>");

        // Apply numbering for the inner list.
        builder.ListFormat.ApplyNumberDefault();

        // Restart numbering for each group and iterate over items.
        builder.Writeln("<<restartNum>><<foreach [item in group.Items]>> <<[item.Name]>> <</foreach>>");

        // Remove numbering after the list.
        builder.ListFormat.RemoveNumbers();

        // End outer foreach.
        builder.Writeln("<</foreach>>");

        // Save the template.
        doc.Save(templatePath);

        // Load the template for reporting.
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();

        // Build the report.
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        var outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}

// Data model classes.
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
}
