using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample XML data.
        string xml = @"<items>
    <item category='Fruits' name='Apple' />
    <item category='Fruits' name='Banana' />
    <item category='Vegetables' name='Carrot' />
    <item category='Fruits' name='Orange' />
    <item category='Vegetables' name='Lettuce' />
</items>";
        string xmlPath = Path.Combine(Directory.GetCurrentDirectory(), "items.xml");
        File.WriteAllText(xmlPath, xml);

        // Load XML and create a grouped data model.
        ReportModel model = new ReportModel
        {
            Groups = LoadAndGroup(xmlPath)
        };

        // -----------------------------------------------------------------
        // 1. Create the template document programmatically.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Outer foreach iterates over groups.
        builder.Writeln("<<foreach [group in model.Groups]>>");
        builder.Writeln("Group: <<[group.Category]>>");
        builder.Writeln();

        // Inner foreach iterates over items inside the current group.
        builder.Writeln("<<foreach [item in group.Items]>>");
        builder.Writeln("- <<[item.Name]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report.
        // -----------------------------------------------------------------
        Document report = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        // No special options are required for this example.
        engine.BuildReport(report, model, "model");

        // Save the generated report.
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        report.Save(reportPath);

        Console.WriteLine("Report generated successfully:");
        Console.WriteLine(reportPath);
    }

    // Loads the XML file, groups items by the 'category' attribute,
    // and returns a list of Group objects ready for the report.
    private static List<Group> LoadAndGroup(string xmlFilePath)
    {
        // Load XML into an XDocument.
        var xdoc = System.Xml.Linq.XDocument.Load(xmlFilePath);

        // Project XML elements into Item objects.
        var items = xdoc.Root!
            .Elements("item")
            .Select(e => new Item
            {
                Category = (string?)e.Attribute("category") ?? string.Empty,
                Name = (string?)e.Attribute("name") ?? string.Empty
            })
            .ToList();

        // Group items by Category.
        var groups = items
            .GroupBy(i => i.Category)
            .Select(g => new Group
            {
                Category = g.Key,
                Items = g.ToList()
            })
            .ToList();

        return groups;
    }
}

// Root data model passed to the reporting engine.
public class ReportModel
{
    public List<Group> Groups { get; set; } = new();
}

// Represents a group of items sharing the same category.
public class Group
{
    public string Category { get; set; } = string.Empty;
    public List<Item> Items { get; set; } = new();
}

// Simple item model.
public class Item
{
    public string Category { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
}
