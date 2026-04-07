using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new ReportModel
        {
            Groups = new()
            {
                new Group
                {
                    Name = "Fruits",
                    Items = new()
                    {
                        new Item { Name = "Apple" },
                        new Item { Name = "Banana" },
                        new Item { Name = "Cherry" }
                    }
                },
                new Group
                {
                    Name = "Vegetables",
                    Items = new()
                    {
                        new Item { Name = "Carrot" },
                        new Item { Name = "Lettuce" },
                        new Item { Name = "Tomato" }
                    }
                }
            }
        };

        // -----------------------------------------------------------------
        // 1. Create the template document programmatically.
        // -----------------------------------------------------------------
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Use a numbered list style for the items.
        builder.ListFormat.List = template.Lists.Add(ListTemplate.NumberDefault);

        // Outer foreach – iterate over groups.
        builder.Writeln("<<foreach [g in Groups]>>");

        // Write the group name as a heading (plain text).
        builder.Writeln("<<[g.Name]>>");

        // Inner foreach – iterate over items.
        // The <<restartNum>> tag is placed before the inner <<foreach>> in the same numbered paragraph.
        builder.Writeln("1. <<restartNum>><<foreach [i in g.Items]>> <<[i.Name]>> <</foreach>>");

        // Close the outer foreach.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report.
        // -----------------------------------------------------------------
        var doc = new Document(templatePath);
        var engine = new ReportingEngine();

        // Build the report using the model as the root data source named "model".
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}

// ---------------------------------------------------------------------
// Data model classes – must be public with public properties.
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
}
