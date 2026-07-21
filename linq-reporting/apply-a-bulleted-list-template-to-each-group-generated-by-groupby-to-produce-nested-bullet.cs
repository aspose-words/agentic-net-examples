using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create the data model with groups and items.
        ReportModel model = new ReportModel
        {
            Groups = new List<Group>
            {
                new Group
                {
                    Category = "Fruits",
                    Items = new List<Item>
                    {
                        new Item { Name = "Apple" },
                        new Item { Name = "Banana" }
                    }
                },
                new Group
                {
                    Category = "Vegetables",
                    Items = new List<Item>
                    {
                        new Item { Name = "Carrot" },
                        new Item { Name = "Lettuce" }
                    }
                }
            }
        };

        // Create a blank document and a builder to construct the template.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define list templates: outer list for groups, inner list for items.
        List outerList = doc.Lists.Add(ListTemplate.BulletDefault);      // •
        List innerList = doc.Lists.Add(ListTemplate.BulletCircle);      // ○

        // Begin the outer foreach over groups.
        builder.Writeln("<<foreach [group in Groups]>>");

        // Outer bullet – group title.
        builder.ListFormat.List = outerList;
        builder.Writeln("<<[group.Category]>>");

        // Switch to inner list for the group's items.
        builder.ListFormat.List = innerList;
        builder.Writeln("<<foreach [item in group.Items]>>");
        builder.Writeln("<<[item.Name]>>");
        builder.Writeln("<</foreach>>");

        // Return to outer list for the next group.
        builder.ListFormat.List = outerList;

        // End the outer foreach.
        builder.Writeln("<</foreach>>");

        // Clean up list formatting.
        builder.ListFormat.RemoveNumbers();

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(doc, model, "model");

        // Save the generated document.
        doc.Save("Report.docx");
    }
}

// Data model classes.
public class ReportModel
{
    public List<Group> Groups { get; set; } = new();
}

public class Group
{
    public string Category { get; set; } = string.Empty;
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = string.Empty;
}
