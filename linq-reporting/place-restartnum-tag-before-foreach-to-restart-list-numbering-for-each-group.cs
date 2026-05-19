using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Reporting;

public class Item
{
    public string Name { get; set; } = "";
}

public class Group
{
    public string Name { get; set; } = "";
    public List<Item> Items { get; set; } = new();
}

public class ReportModel
{
    public List<Group> groups { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Sample data.
        var model = new ReportModel
        {
            groups = new List<Group>
            {
                new Group
                {
                    Name = "Group A",
                    Items = new List<Item>
                    {
                        new Item { Name = "Item 1" },
                        new Item { Name = "Item 2" }
                    }
                },
                new Group
                {
                    Name = "Group B",
                    Items = new List<Item>
                    {
                        new Item { Name = "Item 1" },
                        new Item { Name = "Item 2" },
                        new Item { Name = "Item 3" }
                    }
                }
            }
        };

        // Create the template document.
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Outer foreach over groups.
        builder.Writeln("<<foreach [group in groups]>>");
        // Group heading.
        builder.Writeln("<<[group.Name]>>");

        // Create a numbered list and apply it to the paragraph that will contain the items.
        var list = template.Lists.Add(ListTemplate.NumberDefault);
        builder.ListFormat.List = list;

        // Restart numbering for each group and iterate over items.
        builder.Writeln("1. <<restartNum>><<foreach [item in group.Items]>> <<[item.Name]>>");
        // Close inner and outer foreach tags.
        builder.Writeln("<</foreach>><</foreach>>");
        builder.ListFormat.RemoveNumbers();

        // Save the template, then load it for reporting.
        const string templatePath = "Template.docx";
        template.Save(templatePath);
        var loadedTemplate = new Document(templatePath);

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(loadedTemplate, model);

        // Save the final report.
        loadedTemplate.Save("Report.docx");
    }
}
