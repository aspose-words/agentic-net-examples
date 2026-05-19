using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // 1. Create sample XML data.
        string xmlPath = Path.Combine(outputDir, "Items.xml");
        CreateSampleXml(xmlPath);

        // 2. Load XML and transform into a grouped data model.
        ReportModel model = LoadAndGroupData(xmlPath);

        // 3. Create the LINQ Reporting template.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        CreateTemplateDocument(templatePath);

        // 4. Load the template and build the report.
        Document template = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(template, model, "model");

        // 5. Save the generated report.
        string reportPath = Path.Combine(outputDir, "Report.docx");
        template.Save(reportPath);
    }

    // Creates a simple XML file with items that have a Category attribute.
    private static void CreateSampleXml(string filePath)
    {
        XDocument doc = new XDocument(
            new XElement("Items",
                new XElement("Item", new XAttribute("Category", "Fruits"), new XAttribute("Name", "Apple")),
                new XElement("Item", new XAttribute("Category", "Fruits"), new XAttribute("Name", "Banana")),
                new XElement("Item", new XAttribute("Category", "Vegetables"), new XAttribute("Name", "Carrot")),
                new XElement("Item", new XAttribute("Category", "Fruits"), new XAttribute("Name", "Orange")),
                new XElement("Item", new XAttribute("Category", "Vegetables"), new XAttribute("Name", "Lettuce"))
            )
        );
        doc.Save(filePath);
    }

    // Loads the XML, creates Item objects, groups them by Category, and builds the report model.
    private static ReportModel LoadAndGroupData(string xmlPath)
    {
        XDocument doc = XDocument.Load(xmlPath);
        List<Item> items = doc.Root!
            .Elements("Item")
            .Select(x => new Item
            {
                Category = (string?)x.Attribute("Category") ?? string.Empty,
                Name = (string?)x.Attribute("Name") ?? string.Empty
            })
            .ToList();

        List<ItemGroup> groups = items
            .GroupBy(i => i.Category)
            .Select(g => new ItemGroup
            {
                Category = g.Key,
                Items = g.ToList()
            })
            .ToList();

        return new ReportModel { Groups = groups };
    }

    // Generates a Word document containing LINQ Reporting tags.
    private static void CreateTemplateDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Outer loop over groups.
        builder.Writeln("<<foreach [group in Groups]>>");
        builder.Writeln("Group: <<[group.Category]>>");
        builder.Writeln("");

        // Inner loop over items within the current group.
        builder.Writeln("<<foreach [item in group.Items]>>");
        builder.Writeln("- <<[item.Name]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("");
        builder.Writeln("<</foreach>>");

        doc.Save(filePath);
    }
}

// Root model passed to the reporting engine.
public class ReportModel
{
    public List<ItemGroup> Groups { get; set; } = new();
}

// Represents a group of items sharing the same category.
public class ItemGroup
{
    public string Category { get; set; } = string.Empty;
    public List<Item> Items { get; set; } = new();
}

// Represents a single item.
public class Item
{
    public string Category { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
}
