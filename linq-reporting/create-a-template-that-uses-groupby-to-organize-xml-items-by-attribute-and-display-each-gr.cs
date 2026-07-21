using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

public partial class Program
{
    // Entry point
    public static void Main()
    {
        // 1. Prepare sample XML data
        const string xmlFileName = "items.xml";
        CreateSampleXml(xmlFileName);

        // 2. Load XML and transform it into a grouped data model
        ReportModel model = BuildReportModelFromXml(xmlFileName);

        // 3. Create a LINQ Reporting template programmatically
        const string templateFileName = "template.docx";
        CreateTemplateDocument(templateFileName);

        // 4. Load the template and build the report
        Document report = new Document(templateFileName);
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(report, model, "model");

        // 5. Save the generated report
        const string outputFileName = "report.docx";
        report.Save(outputFileName);
    }

    // Creates a simple XML file with items that have a Category attribute
    private static void CreateSampleXml(string filePath)
    {
        var doc = new XDocument(
            new XElement("Items",
                new XElement("Item", new XAttribute("Category", "Fruit"), new XAttribute("Name", "Apple")),
                new XElement("Item", new XAttribute("Category", "Fruit"), new XAttribute("Name", "Banana")),
                new XElement("Item", new XAttribute("Category", "Vegetable"), new XAttribute("Name", "Carrot")),
                new XElement("Item", new XAttribute("Category", "Vegetable"), new XAttribute("Name", "Lettuce")),
                new XElement("Item", new XAttribute("Category", "Beverage"), new XAttribute("Name", "Coffee"))
            )
        );
        doc.Save(filePath);
    }

    // Reads the XML file, creates Item objects, groups them by Category,
    // and builds the wrapper model required by the reporting engine.
    private static ReportModel BuildReportModelFromXml(string xmlPath)
    {
        var xdoc = XDocument.Load(xmlPath);

        // Parse XML into a flat list of Item objects
        List<Item> items = xdoc.Root!
            .Elements("Item")
            .Select(e => new Item
            {
                Category = (string?)e.Attribute("Category") ?? string.Empty,
                Name = (string?)e.Attribute("Name") ?? string.Empty
            })
            .ToList();

        // Group items by Category
        List<Group> groups = items
            .GroupBy(i => i.Category)
            .Select(g => new Group
            {
                Key = g.Key,
                Items = g.ToList()
            })
            .ToList();

        // Return the model that will be passed to the engine
        return new ReportModel { Groups = groups };
    }

    // Generates a Word document containing LINQ Reporting tags.
    private static void CreateTemplateDocument(string filePath)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Optional title
        builder.Writeln("Items grouped by Category");
        builder.Writeln();

        // Outer foreach iterates over groups
        builder.Writeln("<<foreach [group in Groups]>>");
        builder.Writeln("Category: <<[group.Key]>>");
        builder.Writeln();

        // Inner foreach iterates over items within the current group
        builder.Writeln("<<foreach [item in group.Items]>>");
        builder.Writeln("- <<[item.Name]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln();
        builder.Writeln("<</foreach>>");

        // Save the template for later loading
        doc.Save(filePath);
    }
}

// ---------------------------------------------------------------------------
// Data model classes – must be public with public properties for the engine.
// ---------------------------------------------------------------------------

public class ReportModel
{
    public List<Group> Groups { get; set; } = new();
}

public class Group
{
    public string Key { get; set; } = string.Empty;
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Category { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
}
