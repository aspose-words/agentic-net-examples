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
        const string xmlPath = "data.xml";
        File.WriteAllText(xmlPath,
@"<Categories>
    <Category Name=""Food"">
        <Item Name=""Apple"" Amount=""1.20"" />
        <Item Name=""Bread"" Amount=""2.50"" />
    </Category>
    <Category Name=""Beverages"">
        <Item Name=""Coffee"" Amount=""3.00"" />
        <Item Name=""Tea"" Amount=""2.25"" />
    </Category>
</Categories>");

        // Load XML into strongly‑typed model.
        var model = new ReportModel
        {
            Categories = XDocument.Load(xmlPath)
                .Root!
                .Elements("Category")
                .Select(c => new Category
                {
                    Name = (string)c.Attribute("Name")!,
                    Items = c.Elements("Item")
                             .Select(i => new Item
                             {
                                 Name = (string)i.Attribute("Name")!,
                                 Amount = decimal.Parse((string)i.Attribute("Amount")!)
                             })
                             .ToList()
                })
                .ToList()
        };

        // Create the template document with LINQ Reporting tags.
        const string templatePath = "template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("<<foreach [cat in Categories]>>");
        builder.Writeln("Category: <<[cat.Name]>>");
        builder.Writeln("<<foreach [itm in cat.Items]>>");
        builder.Writeln("- <<[itm.Name]>>: $<<[itm.Amount]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("Total: $<<[cat.Total]>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // Load the template and build the report.
        var doc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        const string outputPath = "report.docx";
        doc.Save(outputPath);
    }
}

// Root data model.
public class ReportModel
{
    public List<Category> Categories { get; set; } = new();
}

// Category with a computed total.
public class Category
{
    public string Name { get; set; } = string.Empty;
    public List<Item> Items { get; set; } = new();

    // Inline arithmetic can also be done in the template,
    // but exposing the total here simplifies the example.
    public decimal Total => Items.Sum(i => i.Amount);
}

// Simple item model.
public class Item
{
    public string Name { get; set; } = string.Empty;
    public decimal Amount { get; set; }
}
