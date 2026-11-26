using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingXmlTotals
{
    public class Item
    {
        public string Name { get; set; } = "";
        public decimal Price { get; set; }
        public int Quantity { get; set; }
    }

    public class Category
    {
        public string Name { get; set; } = "";
        public List<Item> Items { get; set; } = new();
        public decimal Total => Items.Sum(i => i.Price * i.Quantity);
    }

    public class ReportModel
    {
        public List<Category> Categories { get; set; } = new();
    }

    class Program
    {
        static void Main()
        {
            // Register code page provider for XML encoding support.
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            // Create sample XML data.
            string xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Report>
  <Category>
    <Name>Fruits</Name>
    <Item>
      <Name>Apple</Name>
      <Price>0.5</Price>
      <Quantity>10</Quantity>
    </Item>
    <Item>
      <Name>Banana</Name>
      <Price>0.3</Price>
      <Quantity>15</Quantity>
    </Item>
  </Category>
  <Category>
    <Name>Vegetables</Name>
    <Item>
      <Name>Carrot</Name>
      <Price>0.2</Price>
      <Quantity>20</Quantity>
    </Item>
    <Item>
      <Name>Broccoli</Name>
      <Price>0.8</Price>
      <Quantity>5</Quantity>
    </Item>
  </Category>
</Report>";
            string xmlPath = Path.Combine(Directory.GetCurrentDirectory(), "data.xml");
            File.WriteAllText(xmlPath, xmlContent);

            // Load XML into the object model.
            ReportModel model = LoadModelFromXml(xmlPath);

            // Create the LINQ Reporting template.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            builder.Writeln("Categories Report");
            builder.Writeln("<<foreach [cat in Categories]>>");
            builder.Writeln("Category: <<[cat.Name]>>");
            builder.Writeln("<<foreach [itm in cat.Items]>>");
            builder.Writeln("Item: <<[itm.Name]>>, Qty: <<[itm.Quantity]>>, Price: <<[itm.Price]>>, Total: <<[itm.Price * itm.Quantity]>>");
            builder.Writeln("<</foreach>>");
            builder.Writeln("Category Total: <<[cat.Total]>>");
            builder.Writeln("<</foreach>>");

            // Save the template.
            string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
            template.Save(templatePath);

            // Build the report.
            Document reportDoc = new Document(templatePath);
            ReportingEngine engine = new ReportingEngine
            {
                Options = ReportBuildOptions.None
            };
            engine.BuildReport(reportDoc, model, "model");

            // Save the final report.
            string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
            reportDoc.Save(reportPath);
        }

        static ReportModel LoadModelFromXml(string path)
        {
            XDocument xdoc = XDocument.Load(path);
            ReportModel model = new ReportModel();

            foreach (var catElem in xdoc.Root?.Elements("Category") ?? Enumerable.Empty<XElement>())
            {
                Category cat = new Category
                {
                    Name = catElem.Element("Name")?.Value ?? ""
                };

                foreach (var itemElem in catElem.Elements("Item"))
                {
                    Item itm = new Item
                    {
                        Name = itemElem.Element("Name")?.Value ?? "",
                        Price = decimal.TryParse(itemElem.Element("Price")?.Value, out var p) ? p : 0m,
                        Quantity = int.TryParse(itemElem.Element("Quantity")?.Value, out var q) ? q : 0
                    };
                    cat.Items.Add(itm);
                }

                model.Categories.Add(cat);
            }

            return model;
        }
    }
}
