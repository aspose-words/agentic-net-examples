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
        Directory.CreateDirectory("output");

        // 1. Create sample XML data.
        const string xmlPath = "sample.xml";
        var xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Root>
  <Category>
    <Name>Alpha</Name>
    <Item>
      <Title>Item A1</Title>
      <Description>Description for A1</Description>
    </Item>
    <Item>
      <Title>Item A2</Title>
      <Description>Description for A2</Description>
    </Item>
  </Category>
  <Category>
    <Name>Beta</Name>
    <Item>
      <Title>Item B1</Title>
      <Description>Description for B1</Description>
    </Item>
  </Category>
  <Category>
    <Name>Gamma</Name>
    <Item>
      <Title>Item G1</Title>
      <Description>Description for G1</Description>
    </Item>
    <Item>
      <Title>Item G2</Title>
      <Description>Description for G2</Description>
    </Item>
    <Item>
      <Title>Item G3</Title>
      <Description>Description for G3</Description>
    </Item>
  </Category>
</Root>";
        File.WriteAllText(xmlPath, xmlContent);

        // 2. Load XML and filter categories whose name starts with 'A' (case‑insensitive).
        XDocument xDoc = XDocument.Load(xmlPath);
        var filteredCategories = xDoc.Root!
            .Elements("Category")
            .Where(c => ((string?)c.Element("Name"))?.StartsWith("A", StringComparison.OrdinalIgnoreCase) == true)
            .Select(c => new Category
            {
                Name = (string?)c.Element("Name") ?? string.Empty,
                Items = c.Elements("Item")
                         .Select(i => new Item
                         {
                             Title = (string?)i.Element("Title") ?? string.Empty,
                             Description = (string?)i.Element("Description") ?? string.Empty
                         })
                         .ToList()
            })
            .ToList();

        // 3. Prepare the root model.
        var model = new ReportModel { Categories = filteredCategories };

        // 4. Create the LINQ Reporting template programmatically.
        const string templatePath = "output/template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Outer numbered list for categories.
        builder.ListFormat.ApplyNumberDefault();
        builder.Writeln("<<foreach [cat in Model.Categories]>>");
        builder.Writeln("<<restartNum>><<[cat.Name]>>");

        // Inner numbered list for items.
        builder.ListFormat.ListIndent();
        builder.Writeln("<<foreach [itm in cat.Items]>>");
        builder.Writeln("<<restartNum>><<[itm.Title]>> - <<[itm.Description]>>");
        builder.Writeln("<</foreach>>");
        builder.ListFormat.ListOutdent();

        builder.Writeln("<</foreach>>");
        builder.ListFormat.RemoveNumbers();

        // Save the template.
        templateDoc.Save(templatePath);

        // 5. Load the template for report generation.
        Document reportDoc = new Document(templatePath);

        // 6. Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(reportDoc, model, "Model");

        // 7. Save the final report.
        const string reportPath = "output/report.docx";
        reportDoc.Save(reportPath);
    }
}

// Root wrapper class.
public class ReportModel
{
    public List<Category> Categories { get; set; } = new();
}

// Category class.
public class Category
{
    public string Name { get; set; } = string.Empty;
    public List<Item> Items { get; set; } = new();
}

// Item class.
public class Item
{
    public string Title { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
}
