using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Register code page provider for XML parsing.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // -----------------------------------------------------------------
        // Sample XML data.
        // -----------------------------------------------------------------
        string xmlContent = @"
<Root>
  <Product Category=""Electronics"">
    <Name>Phone</Name>
    <Quantity>10</Quantity>
    <Price>299.99</Price>
  </Product>
  <Product Category=""Electronics"">
    <Name>Laptop</Name>
    <Quantity>5</Quantity>
    <Price>799.99</Price>
  </Product>
  <Product Category=""Books"">
    <Name>Novel</Name>
    <Quantity>20</Quantity>
    <Price>9.99</Price>
  </Product>
  <Product Category=""Books"">
    <Name>Comics</Name>
    <Quantity>15</Quantity>
    <Price>4.99</Price>
  </Product>
</Root>";
        const string xmlPath = "Data.xml";
        File.WriteAllText(xmlPath, xmlContent, Encoding.UTF8);

        // -----------------------------------------------------------------
        // Transform XML into a grouped model.
        // -----------------------------------------------------------------
        XDocument xdoc = XDocument.Load(xmlPath);
        List<Group> groups = xdoc.Root!
            .Elements("Product")
            .GroupBy(p => (string)p.Attribute("Category")!)
            .Select(g => new Group
            {
                Category = g.Key,
                Items = g.Select(i => new Item
                {
                    Name = (string)i.Element("Name") ?? string.Empty,
                    Quantity = (int?)i.Element("Quantity") ?? 0,
                    Price = (decimal?)i.Element("Price") ?? 0m
                }).ToList()
            }).ToList();

        var model = new ReportModel { Groups = groups };

        // -----------------------------------------------------------------
        // Build the template document programmatically.
        // -----------------------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Product Report");
        builder.Writeln();

        // Outer foreach – groups.
        builder.Writeln("<<foreach [group in Groups]>>");
        builder.Writeln("Category: <<[group.Category]>>");
        builder.Writeln();

        // Inner foreach – items. The foreach tag is placed BEFORE the table
        // so that no paragraph is written while the builder is inside a table.
        builder.Writeln("<<foreach [item in group.Items]>>");
        Table table = builder.StartTable();

        // Header row (will be repeated for each item – acceptable for this demo).
        builder.InsertCell(); builder.Writeln("Item Name");
        builder.InsertCell(); builder.Writeln("Quantity");
        builder.InsertCell(); builder.Writeln("Price");
        builder.EndRow();

        // Data row.
        builder.InsertCell(); builder.Writeln("<<[item.Name]>>");
        builder.InsertCell(); builder.Writeln("<<[item.Quantity]>>");
        builder.InsertCell(); builder.Writeln("<<[item.Price]>>");
        builder.EndRow();

        builder.EndTable();
        builder.Writeln("<</foreach>>"); // End inner foreach.

        builder.Writeln(); // Blank line between groups.
        builder.Writeln("<</foreach>>"); // End outer foreach.

        // Save the template.
        const string templatePath = "Template.docx";
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Build the final report.
        // -----------------------------------------------------------------
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        const string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}

// ---------------------------------------------------------------------
// Data model.
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
    public string Name { get; set; } = string.Empty;
    public int Quantity { get; set; }
    public decimal Price { get; set; }
}
