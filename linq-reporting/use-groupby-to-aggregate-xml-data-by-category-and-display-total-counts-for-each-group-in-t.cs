using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExample
{
    // Represents a single item from the XML source.
    public class Item
    {
        public string Category { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
    }

    // Holds aggregated data for a category.
    public class CategoryGroup
    {
        public string Category { get; set; } = string.Empty;
        public int Count { get; set; }
    }

    // Root object passed to the ReportingEngine.
    public class ReportModel
    {
        public List<CategoryGroup> Groups { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Create sample XML data.
            // -----------------------------------------------------------------
            const string xmlFileName = "data.xml";
            const string xmlContent = @"
<Items>
    <Item><Category>Fruit</Category><Name>Apple</Name></Item>
    <Item><Category>Fruit</Category><Name>Banana</Name></Item>
    <Item><Category>Vegetable</Category><Name>Carrot</Name></Item>
    <Item><Category>Fruit</Category><Name>Orange</Name></Item>
    <Item><Category>Vegetable</Category><Name>Broccoli</Name></Item>
    <Item><Category>Dairy</Category><Name>Milk</Name></Item>
</Items>";
            File.WriteAllText(xmlFileName, xmlContent.Trim());

            // -----------------------------------------------------------------
            // 2. Load XML and aggregate by Category using LINQ GroupBy.
            // -----------------------------------------------------------------
            XDocument xDoc = XDocument.Load(xmlFileName);
            List<Item> items = xDoc.Descendants("Item")
                .Select(x => new Item
                {
                    Category = (string?)x.Element("Category") ?? string.Empty,
                    Name = (string?)x.Element("Name") ?? string.Empty
                })
                .ToList();

            List<CategoryGroup> groups = items
                .GroupBy(i => i.Category)
                .Select(g => new CategoryGroup
                {
                    Category = g.Key,
                    Count = g.Count()
                })
                .ToList();

            ReportModel model = new ReportModel { Groups = groups };

            // -----------------------------------------------------------------
            // 3. Build a template document programmatically.
            // -----------------------------------------------------------------
            const string templateFileName = "template.docx";
            Document templateDoc = new Document();
            DocumentBuilder builder = new DocumentBuilder(templateDoc);

            builder.Writeln("Category Report");
            builder.Writeln(); // empty line

            // LINQ Reporting foreach loop over Groups.
            builder.Writeln("<<foreach [g in Groups]>>");
            builder.Writeln("Category: <<[g.Category]>> - Total Items: <<[g.Count]>>");
            builder.Writeln("<</foreach>>");

            // Save the template so it can be loaded later.
            templateDoc.Save(templateFileName);

            // -----------------------------------------------------------------
            // 4. Load the template and generate the report.
            // -----------------------------------------------------------------
            Document reportDoc = new Document(templateFileName);
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None; // default options

            // BuildReport with root object name "model" to match the tags.
            engine.BuildReport(reportDoc, model, "model");

            // -----------------------------------------------------------------
            // 5. Save the final report.
            // -----------------------------------------------------------------
            const string outputFileName = "CategoryReport.docx";
            reportDoc.Save(outputFileName);

            // Inform the user (no interactive input required).
            Console.WriteLine($"Report generated: {Path.GetFullPath(outputFileName)}");
        }
    }
}
