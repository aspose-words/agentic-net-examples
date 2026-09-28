using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public string Title { get; set; } = "";
    public string Description { get; set; } = "";
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
        const string xmlFileName = "items.xml";
        var xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Items>
    <Item>
        <Title>First Item</Title>
        <Description>This is the first item.</Description>
    </Item>
    <Item>
        <Title>Second Item</Title>
        <Description>This is the second item.</Description>
    </Item>
    <Item>
        <Title>Third Item</Title>
        <Description>This is the third item.</Description>
    </Item>
</Items>";
        File.WriteAllText(xmlFileName, xmlContent);

        // Load XML into the data model.
        var model = new ReportModel();
        var xdoc = XDocument.Load(xmlFileName);
        foreach (var elem in xdoc.Root?.Elements("Item") ?? [])
        {
            var item = new Item
            {
                Title = (string?)elem.Element("Title") ?? "",
                Description = (string?)elem.Element("Description") ?? ""
            };
            model.Items.Add(item);
        }

        // Create the template document programmatically.
        const string templateFileName = "template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Report generated from XML data");
        builder.Writeln(); // empty line

        // Single‑row template applied to each XML node.
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Title: <<[item.Title]>>");
        builder.Writeln("Description: <<[item.Description]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(templateFileName);

        // Load the template for reporting.
        var templateDoc = new Document(templateFileName);

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(templateDoc, model, "model");

        // Save the final report.
        const string outputFileName = "output.docx";
        templateDoc.Save(outputFileName);
    }
}
