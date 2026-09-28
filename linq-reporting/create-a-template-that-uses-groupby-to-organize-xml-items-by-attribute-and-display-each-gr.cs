using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public string Category { get; set; } = "";
    public string Name { get; set; } = "";
}

public class Group
{
    public string Key { get; set; } = "";
    public List<Item> Items { get; set; } = new();
}

public class ReportModel
{
    public List<Item> Items { get; set; } = new();
    public List<Group> Groups { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare sample XML data.
        string xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Items>
    <Item Category=""Fruits"" Name=""Apple"" />
    <Item Category=""Fruits"" Name=""Banana"" />
    <Item Category=""Vegetables"" Name=""Carrot"" />
    <Item Category=""Fruits"" Name=""Orange"" />
    <Item Category=""Vegetables"" Name=""Lettuce"" />
</Items>";
        string xmlPath = "items.xml";
        File.WriteAllText(xmlPath, xmlContent);

        // Load XML into model.
        ReportModel model = new();
        XDocument doc = XDocument.Load(xmlPath);
        foreach (XElement elem in doc.Root!.Elements("Item"))
        {
            model.Items.Add(new Item
            {
                Category = (string?)elem.Attribute("Category") ?? "",
                Name = (string?)elem.Attribute("Name") ?? ""
            });
        }

        // Create groups for reporting (Category -> Items).
        model.Groups = model.Items
            .GroupBy(i => i.Category)
            .Select(g => new Group
            {
                Key = g.Key,
                Items = g.ToList()
            })
            .ToList();

        // Create template document.
        Document template = new();
        DocumentBuilder builder = new(template);

        builder.Writeln("Items grouped by Category:");
        builder.Writeln("<<foreach [g in Groups]>>");
        builder.Writeln("Category: <<[g.Key]>>");
        builder.Writeln("<<foreach [i in g.Items]>>");
        builder.Writeln("- <<[i.Name]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</foreach>>");

        string templatePath = "template.docx";
        template.Save(templatePath);

        // Load template for reporting.
        Document reportDoc = new(templatePath);
        ReportingEngine engine = new();

        // Build the report.
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        string outputPath = "output.docx";
        reportDoc.Save(outputPath);
    }
}
