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
        // Ensure code page provider is registered (required for some Aspose.Words features)
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // 1. Create sample XML data
        const string xmlFile = "data.xml";
        File.WriteAllText(xmlFile,
@"<Report>
  <Categories>
    <Category Name='Beverages'>
      <Item Name='Coffee' Price='3.5' Quantity='10' />
      <Item Name='Tea' Price='2.0' Quantity='15' />
    </Category>
    <Category Name='Snacks'>
      <Item Name='Cookie' Price='1.2' Quantity='20' />
      <Item Name='Chips' Price='1.5' Quantity='12' />
    </Category>
  </Categories>
</Report>");

        // 2. Load XML into model objects
        XDocument doc = XDocument.Load(xmlFile);
        ReportModel model = new()
        {
            Categories = doc.Root!
                .Element("Categories")!
                .Elements("Category")
                .Select(cat => new Category
                {
                    Name = (string)cat.Attribute("Name")!,
                    Items = cat.Elements("Item")
                               .Select(it => new Item
                               {
                                   Name = (string)it.Attribute("Name")!,
                                   Price = decimal.Parse((string)it.Attribute("Price")!),
                                   Quantity = int.Parse((string)it.Attribute("Quantity")!)
                               })
                               .ToList()
                })
                .ToList()
        };

        // 3. Create a Word template with LINQ Reporting tags
        const string templateFile = "template.docx";
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);

        builder.Writeln("Categories Report");
        builder.Writeln("");

        // Iterate over categories
        builder.Writeln("<<foreach [cat in Categories]>>");
        builder.Writeln("Category: <<[cat.Name]>>");
        // Inline arithmetic using LINQ Sum to calculate total per category
        builder.Writeln("Total: <<[cat.Items.Sum(i => i.Price * i.Quantity)]>>");
        builder.Writeln("");

        // Iterate over items within the current category
        builder.Writeln("<<foreach [item in cat.Items]>>");
        builder.Writeln("- <<[item.Name]>>: <<[item.Price]>> x <<[item.Quantity]>> = <<[item.Price * item.Quantity]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</foreach>>");

        // Save the template
        templateDoc.Save(templateFile);

        // 4. Build the report using Aspose.Words ReportingEngine
        Document reportDoc = new(templateFile);
        ReportingEngine engine = new();
        engine.BuildReport(reportDoc, model, "model");

        // 5. Save the generated report
        const string outputFile = "report.docx";
        reportDoc.Save(outputFile);

        Console.WriteLine($"Report generated: {Path.GetFullPath(outputFile)}");
    }
}

// Data model classes
public class ReportModel
{
    public List<Category> Categories { get; set; } = new();
}

public class Category
{
    public string Name { get; set; } = string.Empty;
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = string.Empty;
    public decimal Price { get; set; }
    public int Quantity { get; set; }
}
