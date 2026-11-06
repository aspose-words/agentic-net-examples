using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public int Id { get; set; } = 0;
    public string Name { get; set; } = "";
    public string Category { get; set; } = "";
}

public class ItemGroup
{
    public string Key { get; set; } = "";
    public List<Item> Items { get; set; } = new();
}

public class ReportModel
{
    public List<ItemGroup> Groups { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare sample XML data.
        string xmlContent = @"
<Items>
    <Item Id='1' Name='Apple' Category='Fruit' />
    <Item Id='2' Name='Banana' Category='Fruit' />
    <Item Id='3' Name='Carrot' Category='Vegetable' />
    <Item Id='4' Name='Broccoli' Category='Vegetable' />
    <Item Id='5' Name='Chicken' Category='Meat' />
</Items>";
        XDocument xDoc = XDocument.Parse(xmlContent);
        List<Item> items = xDoc.Root!
            .Elements("Item")
            .Select(e => new Item
            {
                Id = (int)e.Attribute("Id")!,
                Name = (string)e.Attribute("Name")!,
                Category = (string)e.Attribute("Category")!
            })
            .ToList();

        // Group items by Category.
        var groups = items
            .GroupBy(i => i.Category)
            .Select(g => new ItemGroup
            {
                Key = g.Key,
                Items = g.ToList()
            })
            .ToList();

        // Build the model.
        ReportModel model = new()
        {
            Groups = groups
        };

        // Create the template document.
        string templatePath = "template.docx";
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);

        builder.Writeln("Items grouped by Category:");
        builder.Writeln("<<foreach [group in Groups]>>");
        builder.Writeln("Category: <<[group.Key]>>");
        builder.Writeln("<<foreach [item in group.Items]>>");
        builder.Writeln("- Id: <<[item.Id]>>, Name: <<[item.Name]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // Load the template and generate the report.
        Document doc = new(templatePath);
        ReportingEngine engine = new();
        engine.BuildReport(doc, model, "model");

        // Save the final report.
        string outputPath = "output.docx";
        doc.Save(outputPath);
    }
}
