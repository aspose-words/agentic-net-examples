using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public string Name { get; set; } = "";
    public string Category { get; set; } = "";
    public string Composite => $"{Name} - {Category}";
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
        string xmlContent = @"<?xml version=""1.0"" encoding=""utf-8""?>
<Items>
    <Item Name=""Apple"" Category=""Fruit"" />
    <Item Name=""Carrot"" Category=""Vegetable"" />
    <Item Name=""Banana"" Category=""Fruit"" />
</Items>";
        string xmlPath = "data.xml";
        File.WriteAllText(xmlPath, xmlContent);

        // Load XML and convert to model objects.
        XDocument xDoc = XDocument.Load(xmlPath);
        var model = new ReportModel();
        foreach (XElement elem in xDoc.Root?.Elements("Item") ?? new List<XElement>())
        {
            var item = new Item
            {
                Name = (string?)elem.Attribute("Name") ?? "",
                Category = (string?)elem.Attribute("Category") ?? ""
            };
            model.Items.Add(item);
        }

        // Create the LINQ Reporting template.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("Item List:");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("<<[item.Composite]>>");
        builder.Writeln("<</foreach>>");
        string templatePath = "template.docx";
        template.Save(templatePath);

        // Load the template for reporting.
        Document doc = new Document(templatePath);

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string outputPath = "output.docx";
        doc.Save(outputPath);
    }
}
