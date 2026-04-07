using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample XML data.
        const string xmlFileName = "data.xml";
        var xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Root>
    <Item>
        <Name>Apple</Name>
        <Value>10</Value>
    </Item>
    <Item>
        <Name>Banana</Name>
        <Value>20</Value>
    </Item>
    <Item>
        <Name>Cherry</Name>
        <Value>30</Value>
    </Item>
</Root>";
        File.WriteAllText(xmlFileName, xmlContent);

        // Load XML and convert to a strongly‑typed model.
        var model = new ReportModel();
        var doc = XDocument.Load(xmlFileName);
        foreach (var elem in doc.Root?.Elements("Item") ?? [])
        {
            var item = new Item
            {
                Name = elem.Element("Name")?.Value ?? string.Empty,
                Value = elem.Element("Value")?.Value ?? string.Empty
            };
            model.Items.Add(item);
        }

        // Create a Word template programmatically.
        const string templateFileName = "template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("XML Data Report");

        // Header table (static part).
        builder.StartTable();
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Value");
        builder.EndRow();
        builder.EndTable();

        // Data rows – LINQ Reporting foreach block.
        builder.Writeln("<<foreach [item in model.Items]>>");
        builder.StartTable();
        builder.InsertCell();
        builder.Writeln("<<[item.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[item.Value]>>");
        builder.EndRow();
        builder.EndTable();
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templateFileName);

        // Load the template and build the report.
        var reportDoc = new Document(templateFileName);
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // default options
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        const string outputFileName = "output.docx";
        reportDoc.Save(outputFileName);
    }
}

// Public data model classes.
public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = string.Empty;
    public string Value { get; set; } = string.Empty;
}
