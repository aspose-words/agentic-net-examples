using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using System.Text;

namespace LinqReportingCaseInsensitiveFilter
{
    // Model representing a single item.
    public class Item
    {
        public string Name { get; set; } = "";
        public string Category { get; set; } = "";
    }

    // Root model passed to the reporting engine.
    public class ReportModel
    {
        public List<Item> Items { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required by Aspose.Words for some encodings).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // 1. Create sample XML data.
            string xmlContent = @"
                <Items>
                    <Item Name='Apple' Category='Fruit' />
                    <Item Name='Carrot' Category='Vegetable' />
                    <Item Name='Banana' Category='fruit' />
                    <Item Name='Broccoli' Category='Vegetable' />
                    <Item Name='Strawberry' Category='FRUIT' />
                </Items>";
            string xmlPath = Path.Combine(Directory.GetCurrentDirectory(), "Data.xml");
            File.WriteAllText(xmlPath, xmlContent);

            // 2. Load XML and filter nodes with case‑insensitive comparison on the 'Category' attribute.
            XDocument xDoc = XDocument.Load(xmlPath);
            List<Item> filteredItems = xDoc.Root!
                .Elements("Item")
                .Where(e => string.Equals(
                    (string?)e.Attribute("Category"),
                    "fruit",
                    StringComparison.OrdinalIgnoreCase))
                .Select(e => new Item
                {
                    Name = (string?)e.Attribute("Name") ?? "",
                    Category = (string?)e.Attribute("Category") ?? ""
                })
                .ToList();

            // 3. Prepare the report model.
            ReportModel model = new()
            {
                Items = filteredItems
            };

            // 4. Create a Word template with LINQ Reporting tags.
            string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
            Document templateDoc = new();
            DocumentBuilder builder = new(templateDoc);

            builder.Writeln("Filtered Items (Category = 'Fruit'):");
            builder.Writeln("<<foreach [item in Items]>>");
            builder.Writeln(" - <<[item.Name]>> (<<[item.Category]>>)");
            builder.Writeln("<</foreach>>");

            templateDoc.Save(templatePath);

            // 5. Load the template for report generation.
            Document reportDoc = new(templatePath);

            // 6. Build the report.
            ReportingEngine engine = new();
            engine.BuildReport(reportDoc, model, "model");

            // 7. Save the generated report.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
            reportDoc.Save(outputPath);
        }
    }
}
