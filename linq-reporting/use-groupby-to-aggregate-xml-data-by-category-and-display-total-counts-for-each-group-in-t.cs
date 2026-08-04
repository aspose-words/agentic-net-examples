using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingGroupByExample
{
    // Data entity representing a single item from the XML.
    public class Item
    {
        public string Category { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
    }

    // Data entity representing a grouped result (category + total count).
    public class CategoryGroup
    {
        public string Category { get; set; } = string.Empty;
        public int Count { get; set; }
    }

    // Wrapper model that will be passed to the reporting engine.
    public class ReportModel
    {
        public List<CategoryGroup> Groups { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Create sample XML data.
            const string xmlFileName = "SampleData.xml";
            CreateSampleXml(xmlFileName);

            // 2. Load XML and project to a list of Item objects.
            List<Item> items = LoadItemsFromXml(xmlFileName);

            // 3. Group items by Category and calculate total counts.
            List<CategoryGroup> groups = items
                .GroupBy(i => i.Category)
                .Select(g => new CategoryGroup { Category = g.Key, Count = g.Count() })
                .ToList();

            // 4. Prepare the model for the reporting engine.
            ReportModel model = new ReportModel { Groups = groups };

            // 5. Build the template document programmatically.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            builder.Writeln("Category Report");
            builder.Writeln(); // empty line

            // LINQ Reporting tags.
            builder.Writeln("<<foreach [group in Groups]>>");
            builder.Writeln("Category: <<[group.Category]>> - Total Items: <<[group.Count]>>");
            builder.Writeln("<</foreach>>");

            // 6. Generate the final report.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(template, model, "model");

            // 7. Save the output document.
            const string outputFileName = "CategoryReport.docx";
            template.Save(outputFileName);
        }

        // Creates a simple XML file with items belonging to different categories.
        private static void CreateSampleXml(string path)
        {
            XDocument doc = new XDocument(
                new XElement("Items",
                    new XElement("Item",
                        new XElement("Category", "Fruit"),
                        new XElement("Name", "Apple")),
                    new XElement("Item",
                        new XElement("Category", "Fruit"),
                        new XElement("Name", "Banana")),
                    new XElement("Item",
                        new XElement("Category", "Vegetable"),
                        new XElement("Name", "Carrot")),
                    new XElement("Item",
                        new XElement("Category", "Vegetable"),
                        new XElement("Name", "Broccoli")),
                    new XElement("Item",
                        new XElement("Category", "Beverage"),
                        new XElement("Name", "Coffee"))
                ));

            doc.Save(path);
        }

        // Parses the XML file into a list of Item objects.
        private static List<Item> LoadItemsFromXml(string path)
        {
            XDocument doc = XDocument.Load(path);
            return doc.Root?
                .Elements("Item")
                .Select(x => new Item
                {
                    Category = (string?)x.Element("Category") ?? string.Empty,
                    Name = (string?)x.Element("Name") ?? string.Empty
                })
                .ToList() ?? new List<Item>();
        }
    }
}
