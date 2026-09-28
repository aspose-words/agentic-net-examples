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
        // Ensure code page provider is available (required for some environments)
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // 1. Create sample XML data file.
        const string xmlPath = "data.xml";
        var xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Items>
    <Item>
        <Category>Food</Category>
        <Amount>12.50</Amount>
    </Item>
    <Item>
        <Category>Travel</Category>
        <Amount>45.00</Amount>
    </Item>
    <Item>
        <Category>Food</Category>
        <Amount>8.75</Amount>
    </Item>
    <Item>
        <Category>Office</Category>
        <Amount>23.40</Amount>
    </Item>
    <Item>
        <Category>Travel</Category>
        <Amount>15.30</Amount>
    </Item>
</Items>";
        File.WriteAllText(xmlPath, xmlContent);

        // 2. Load XML and map to model.
        XDocument doc = XDocument.Load(xmlPath);
        var model = new ReportModel
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

        // 3. Create the LINQ Reporting template programmatically.
        const string templatePath = "template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Summary by Category:");
        // Begin foreach over grouped items.
        builder.Writeln("<<foreach [g in Items.GroupBy(i => i.Category)]>>");
        builder.Writeln("Category: <<[g.Key]>>");
        builder.Writeln("Total Amount: <<[g.Sum(i => i.Amount)]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // 4. Load the template for report generation.
        var reportDoc = new Document(templatePath);

        // 5. Build the report using ReportingEngine.
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // 6. Save the generated report.
        const string outputPath = "report.docx";
        reportDoc.Save(outputPath);
    }
}
