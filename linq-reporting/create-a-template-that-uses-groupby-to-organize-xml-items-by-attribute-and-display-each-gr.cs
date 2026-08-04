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
        // Prepare sample XML data.
        string xmlContent = @"
<Items>
    <Item Category='Fruits' Name='Apple' />
    <Item Category='Fruits' Name='Banana' />
    <Item Category='Vegetables' Name='Carrot' />
    <Item Category='Fruits' Name='Orange' />
    <Item Category='Vegetables' Name='Lettuce' />
</Items>";

        // Load XML and transform it into grouped model objects.
        XDocument xDoc = XDocument.Parse(xmlContent);
        var groups = xDoc.Root!
            .Elements("Item")
            .Select(e => new Item
            {
                Category = (string?)e.Attribute("Category") ?? string.Empty,
                Name = (string?)e.Attribute("Name") ?? string.Empty
            })
            .GroupBy(i => i.Category)
            .Select(g => new Group
            {
                Category = g.Key,
                Items = g.ToList()
            })
            .ToList();

        var model = new ReportModel { Groups = groups };

        // -----------------------------------------------------------------
        // Create the LINQ Reporting template programmatically.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("LINQ Reporting – Grouped Items");
        builder.Writeln();

        // Outer loop over groups.
        builder.Writeln("<<foreach [group in Groups]>>");
        builder.Writeln("Group: <<[group.Category]>>");
        builder.Writeln();

        // Inner loop over items within the current group.
        builder.Writeln("<<foreach [item in group.Items]>>");
        builder.Writeln("- <<[item.Name]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln();
        builder.Writeln("<</foreach>>");

        // Save the template to disk (required by the lifecycle rule).
        string templatePath = Path.Combine(Environment.CurrentDirectory, "Template.docx");
        template.Save(templatePath);

        // Load the template back (simulating a real scenario where the template exists on disk).
        Document loadedTemplate = new Document(templatePath);

        // -----------------------------------------------------------------
        // Build the report using the ReportingEngine.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // No special flags needed.

        bool success = engine.BuildReport(loadedTemplate, model, "model");

        // Save the generated report.
        string reportPath = Path.Combine(Environment.CurrentDirectory, "Report.docx");
        loadedTemplate.Save(reportPath);

        // The example finishes here; no interactive input is required.
    }
}

// ---------------------------------------------------------------------
// Data model classes used by the template.
// ---------------------------------------------------------------------
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
    public string Category { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
}
