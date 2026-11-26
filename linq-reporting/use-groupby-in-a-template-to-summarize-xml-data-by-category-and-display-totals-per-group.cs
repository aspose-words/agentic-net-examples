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

public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare sample XML data.
        string xmlPath = "Data.xml";
        File.WriteAllText(xmlPath,
@"<Items>
    <Item><Category>Food</Category><Amount>12.5</Amount></Item>
    <Item><Category>Food</Category><Amount>7.3</Amount></Item>
    <Item><Category>Books</Category><Amount>15.0</Amount></Item>
    <Item><Category>Books</Category><Amount>9.99</Amount></Item>
    <Item><Category>Electronics</Category><Amount>199.99</Amount></Item>
</Items>");

        // Load XML into model.
        XDocument doc = XDocument.Load(xmlPath);
        ReportModel model = new()
        {
            Items = doc.Root!
                .Elements("Item")
                .Select(x => new Item
                {
                    Category = (string?)x.Element("Category") ?? "",
                    Amount = decimal.TryParse((string?)x.Element("Amount"), out var v) ? v : 0m
                })
                .ToList()
        };

        // Create template document.
        string templatePath = "Template.docx";
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);

        builder.Writeln("Category Summary Report");
        builder.Writeln("<<foreach [g in Items.GroupBy(i => i.Category)]>>");
        builder.Writeln("Category: <<[g.Key]>>   Total Amount: <<[g.Sum(i => i.Amount)]>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // Load template for reporting.
        Document loadDoc = new(templatePath);

        // Build report.
        ReportingEngine engine = new();
        engine.BuildReport(loadDoc, model, "model");

        // Save final report.
        string outputPath = "Report.docx";
        loadDoc.Save(outputPath);
    }
}
