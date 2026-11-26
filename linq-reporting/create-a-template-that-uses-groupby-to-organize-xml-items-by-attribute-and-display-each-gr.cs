using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public string Name { get; set; } = "";
    public string Category { get; set; } = "";
}

public class Group
{
    public string Category { get; set; } = "";
    public List<Item> Items { get; set; } = new();
}

public class ReportModel
{
    // Root collection of groups.
    public List<Group> Groups { get; set; } = new();

    // Current date/time for the report header.
    public DateTime ReportDate { get; set; } = DateTime.Now;
}

public class Program
{
    public static void Main()
    {
        // ---------- Sample XML data ----------
        string xmlContent = @"
<Items>
    <Item Category='Fruits' Name='Apple' />
    <Item Category='Fruits' Name='Banana' />
    <Item Category='Vegetables' Name='Carrot' />
    <Item Category='Fruits' Name='Orange' />
    <Item Category='Vegetables' Name='Lettuce' />
</Items>";

        // Parse XML into Item objects.
        XDocument xDoc = XDocument.Parse(xmlContent);
        List<Item> items = xDoc.Root!
            .Elements("Item")
            .Select(e => new Item
            {
                Category = (string?)e.Attribute("Category") ?? "",
                Name = (string?)e.Attribute("Name") ?? ""
            })
            .ToList();

        // Group items by Category.
        ReportModel model = new ReportModel
        {
            Groups = items
                .GroupBy(i => i.Category)
                .Select(g => new Group
                {
                    Category = g.Key,
                    Items = g.ToList()
                })
                .ToList()
        };

        // ---------- Create template document ----------
        string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Header with current date.
        builder.Writeln("Report generated on: <<[ReportDate]>>");
        builder.Writeln();

        // Group header.
        builder.Writeln("<<foreach [group in Groups]>>");
        builder.Writeln("Category: <<[group.Category]>>");

        // Items within each group.
        builder.Writeln("<<foreach [item in group.Items]>>");
        builder.Writeln("- <<[item.Name]>>");
        builder.Writeln("<</foreach>>"); // end inner foreach
        builder.Writeln("<</foreach>>"); // end outer foreach

        // Save the template.
        templateDoc.Save(templatePath);

        // ---------- Load template and build report ----------
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();

        // Build the report using the model as the root data source named "model".
        engine.BuildReport(reportDoc, model, "model");

        // Save the final report.
        string outputPath = "report.docx";
        reportDoc.Save(outputPath);
    }
}
