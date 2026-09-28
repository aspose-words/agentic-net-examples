using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingXmlTotals
{
    // Data model classes
    public class ReportModel
    {
        public List<Category> Categories { get; set; } = new();
    }

    public class Category
    {
        public string Name { get; set; } = "";
        public List<Item> Items { get; set; } = new();
    }

    public class Item
    {
        public string Name { get; set; } = "";
        public decimal Price { get; set; }
        public int Quantity { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare sample XML data
            string xmlPath = "data.xml";
            File.WriteAllText(xmlPath,
@"<Report>
  <Categories>
    <Category Name=""Fruits"">
      <Item Name=""Apple"" Price=""1.2"" Quantity=""10""/>
      <Item Name=""Banana"" Price=""0.8"" Quantity=""5""/>
    </Category>
    <Category Name=""Vegetables"">
      <Item Name=""Carrot"" Price=""0.5"" Quantity=""20""/>
      <Item Name=""Tomato"" Price=""0.9"" Quantity=""8""/>
    </Category>
  </Categories>
</Report>");

            // Load XML into the data model
            XDocument xdoc = XDocument.Load(xmlPath);
            ReportModel model = new()
            {
                Categories = xdoc.Root!
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

            // Create the template document programmatically
            string templatePath = "Template.docx";
            Document templateDoc = new();
            DocumentBuilder builder = new(templateDoc);

            builder.Writeln("<<foreach [cat in Categories]>>");
            builder.Writeln("Category: <<[cat.Name]>>");
            builder.Writeln("<<foreach [item in cat.Items]>>");
            builder.Writeln("- <<[item.Name]>>: <<[item.Price]>> x <<[item.Quantity]>> = <<[item.Price * item.Quantity]>>");
            builder.Writeln("<</foreach>>");
            builder.Writeln("Total for <<[cat.Name]>>: <<[cat.Items.Sum(i => i.Price * i.Quantity)]>>");
            builder.Writeln("<</foreach>>");

            templateDoc.Save(templatePath);

            // Load the template for reporting
            Document doc = new(templatePath);

            // Build the report
            ReportingEngine engine = new();
            engine.BuildReport(doc, model, "model");

            // Save the final report
            string outputPath = "ReportOutput.docx";
            doc.Save(outputPath);
        }
    }
}
