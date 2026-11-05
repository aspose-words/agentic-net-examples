using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

#nullable enable

public class Item
{
    public string Name { get; set; } = "";
    public decimal Price { get; set; }
    public int Quantity { get; set; }
}

public class CategoryGroup
{
    public string CategoryName { get; set; } = "";
    public List<Item> Items { get; set; } = new();

    // Calculated total for the group
    public decimal Total => Items.Sum(i => i.Price * i.Quantity);
}

public class ReportModel
{
    public List<CategoryGroup> Groups { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // 1. Create sample XML data
        string xmlPath = "data.xml";
        var xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Orders>
  <Order>
    <Category>Electronics</Category>
    <Name>Smartphone</Name>
    <Price>699.99</Price>
    <Quantity>2</Quantity>
  </Order>
  <Order>
    <Category>Electronics</Category>
    <Name>Laptop</Name>
    <Price>1199.50</Price>
    <Quantity>1</Quantity>
  </Order>
  <Order>
    <Category>Books</Category>
    <Name>Programming C#</Name>
    <Price>49.95</Price>
    <Quantity>3</Quantity>
  </Order>
  <Order>
    <Category>Books</Category>
    <Name>Design Patterns</Name>
    <Price>39.99</Price>
    <Quantity>2</Quantity>
  </Order>
  <Order>
    <Category>Kitchen</Category>
    <Name>Blender</Name>
    <Price>89.99</Price>
    <Quantity>1</Quantity>
  </Order>
</Orders>";
        File.WriteAllText(xmlPath, xmlContent);

        // 2. Load XML and transform into model with grouping
        XDocument doc = XDocument.Load(xmlPath);
        var items = doc.Root!
            .Elements("Order")
            .Select(o => new Item
            {
                Name = (string)o.Element("Name")!,
                Price = decimal.Parse((string)o.Element("Price")!, CultureInfo.InvariantCulture),
                Quantity = int.Parse((string)o.Element("Quantity")!)
            })
            .ToList();

        var groups = items
            .GroupBy(i => (string)doc.Root!
                .Elements("Order")
                .First(o => (string)o.Element("Name")! == i.Name)
                .Element("Category")!)
            .Select(g => new CategoryGroup
            {
                CategoryName = g.Key,
                Items = g.ToList()
            })
            .ToList();

        var model = new ReportModel { Groups = groups };

        // 3. Create Word template with LINQ Reporting tags
        string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Sales Report");
        builder.Writeln("==============");
        builder.Writeln();

        // Iterate over groups
        builder.Writeln("<<foreach [group in Groups]>>");
        builder.Writeln("Category: <<[group.CategoryName]>>");
        builder.Writeln("Group Total: <<[group.Total]>>");
        builder.Writeln();

        // Iterate over items in the current group
        builder.Writeln("<<foreach [item in group.Items]>>");
        Table table = builder.StartTable();

        // Header row
        builder.InsertCell(); builder.Writeln("Item");
        builder.InsertCell(); builder.Writeln("Price");
        builder.InsertCell(); builder.Writeln("Quantity");
        builder.InsertCell(); builder.Writeln("Line Total");
        builder.EndRow();

        // Data row (repeated for each item)
        builder.InsertCell(); builder.Writeln("<<[item.Name]>>");
        builder.InsertCell(); builder.Writeln("<<[item.Price]>>");
        builder.InsertCell(); builder.Writeln("<<[item.Quantity]>>");
        builder.InsertCell(); builder.Writeln("<<[item.Price * item.Quantity]>>");
        builder.EndRow();

        builder.EndTable();
        builder.Writeln("<</foreach>>"); // end item foreach
        builder.Writeln("<</foreach>>"); // end group foreach

        // Save the template
        templateDoc.Save(templatePath);

        // 4. Build the report
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // 5. Save the final report
        string outputPath = "report.docx";
        reportDoc.Save(outputPath);
    }
}
