using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public string Category { get; set; } = "";
    public decimal Amount { get; set; }
}

public class ReportData
{
    public List<Item> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Ensure the working directories exist.
        string outputDir = "output";
        Directory.CreateDirectory(outputDir);

        // 1. Create sample XML data.
        string xmlPath = Path.Combine(outputDir, "data.xml");
        File.WriteAllText(xmlPath,
@"<Items>
    <Item><Category>Food</Category><Amount>12.5</Amount></Item>
    <Item><Category>Food</Category><Amount>7.3</Amount></Item>
    <Item><Category>Books</Category><Amount>15.0</Amount></Item>
    <Item><Category>Books</Category><Amount>9.99</Amount></Item>
    <Item><Category>Electronics</Category><Amount>199.99</Amount></Item>
</Items>");

        // 2. Load XML into objects.
        XDocument xdoc = XDocument.Load(xmlPath);
        var items = xdoc.Root!
            .Elements("Item")
            .Select(x => new Item
            {
                Category = (string?)x.Element("Category") ?? "",
                Amount = decimal.TryParse((string?)x.Element("Amount"), out var a) ? a : 0m
            })
            .ToList();

        var data = new ReportData { Items = items };

        // 3. Create the LINQ Reporting template.
        string templatePath = Path.Combine(outputDir, "template.docx");
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Category Summary Report");
        builder.Writeln("<<foreach [g in data.Items.GroupBy(i => i.Category)]>>");
        builder.Writeln("Category: <<[g.Key]>>");
        builder.Writeln("Total: <<[g.Sum(i => i.Amount)]>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // 4. Load the template and build the report.
        var reportDoc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, data, "data");

        // 5. Save the generated report.
        string resultPath = Path.Combine(outputDir, "output.docx");
        reportDoc.Save(resultPath);
    }
}
