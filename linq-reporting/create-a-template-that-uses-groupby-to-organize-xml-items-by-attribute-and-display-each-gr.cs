using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Item
{
    public string Category { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
}

public class Model
{
    public List<Item> Items { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Create sample XML data.
        string xmlPath = Path.Combine(outputDir, "items.xml");
        string xmlContent =
@"<Items>
    <Item Category=""Fruit"" Name=""Apple"" />
    <Item Category=""Fruit"" Name=""Banana"" />
    <Item Category=""Vegetable"" Name=""Carrot"" />
    <Item Category=""Vegetable"" Name=""Lettuce"" />
    <Item Category=""Fruit"" Name=""Orange"" />
</Items>";
        File.WriteAllText(xmlPath, xmlContent);

        // Load XML into a strongly‑typed model.
        var model = new Model();
        XDocument xdoc = XDocument.Load(xmlPath);
        foreach (var elem in xdoc.Root?.Elements("Item") ?? [])
        {
            model.Items.Add(new Item
            {
                Category = (string?)elem.Attribute("Category") ?? string.Empty,
                Name = (string?)elem.Attribute("Name") ?? string.Empty
            });
        }

        // Create the template document with LINQ Reporting tags.
        string templatePath = Path.Combine(outputDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Items grouped by Category:");
        builder.Writeln("<<foreach [g in Items.GroupBy(i => i.Category)]>>");
        builder.Writeln("Category: <<[g.Key]>>");
        builder.Writeln("<<foreach [item in g]>>");
        builder.Writeln("- <<[item.Name]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // Build the report using the strongly‑typed model as the data source.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model);

        // Save the generated report.
        string resultPath = Path.Combine(outputDir, "report.docx");
        reportDoc.Save(resultPath);

        Console.WriteLine("Report generated at: " + resultPath);
    }
}
