using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Category
{
    public string Name { get; set; } = string.Empty;
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Value { get; set; } = string.Empty;
}

public class ReportModel
{
    public List<Category> Categories { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // 1. Create sample XML data.
        const string xmlFileName = "sample.xml";
        var xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Root>
  <Category Name=""Fruits"">
    <Item>Apple</Item>
    <Item>Banana</Item>
  </Category>
  <Category Name=""Vegetables"">
    <Item>Carrot</Item>
    <Item>Broccoli</Item>
  </Category>
  <Category Name=""Fish"">
    <Item>Salmon</Item>
    <Item>Tuna</Item>
  </Category>
</Root>";
        File.WriteAllText(xmlFileName, xmlContent);

        // 2. Load XML and filter categories whose name starts with 'F'.
        XDocument doc = XDocument.Load(xmlFileName);
        var filteredCategories = doc.Root?
            .Elements("Category")
            .Where(c => ((string?)c.Attribute("Name"))?.StartsWith("F") ?? false)
            .Select(c => new Category
            {
                Name = (string?)c.Attribute("Name") ?? string.Empty,
                Items = c.Elements("Item")
                         .Select(i => new Item { Value = (string?)i ?? string.Empty })
                         .ToList()
            })
            .ToList() ?? new List<Category>();

        var model = new ReportModel { Categories = filteredCategories };

        // 3. Build the LINQ Reporting template programmatically.
        const string templateFileName = "template.docx";
        var docTemplate = new Document();
        var builder = new DocumentBuilder(docTemplate);

        // Title
        builder.Writeln("Filtered Categories (Names starting with 'F'):");
        builder.Writeln();

        // Begin outer foreach over Categories.
        builder.Writeln("<<foreach [cat in Categories]>>");

        // Category bullet (level 0)
        builder.ListFormat.ApplyBulletDefault();
        builder.ListFormat.ListLevelNumber = 0;
        builder.Writeln("<<[cat.Name]>>");

        // Items bullet (level 1)
        builder.ListFormat.ListLevelNumber = 1;
        builder.Writeln("<<foreach [it in cat.Items]>>");
        builder.Writeln("<<[it.Value]>>");
        builder.Writeln("<</foreach>>");

        // Reset list level after inner loop.
        builder.ListFormat.ListLevelNumber = 0;

        // End outer foreach.
        builder.Writeln("<</foreach>>");

        // Save the template.
        docTemplate.Save(templateFileName);

        // 4. Load the template and generate the report.
        var reportDoc = new Document(templateFileName);
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // 5. Save the final report.
        const string outputFileName = "report.docx";
        reportDoc.Save(outputFileName);

        Console.WriteLine($"Report generated: {Path.GetFullPath(outputFileName)}");
    }
}
