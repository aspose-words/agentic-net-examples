using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample XML data
        const string xmlPath = "data.xml";
        GenerateSampleXml(xmlPath, 500);

        // Load data with progress reporting
        var model = new ReportModel();
        model.Items = LoadItemsWithProgress(xmlPath, percent => Console.WriteLine($"Loading XML: {percent}%"));

        // Create template document
        const string templatePath = "template.docx";
        CreateTemplate(templatePath);

        // Build report
        var doc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save output
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
        Console.WriteLine($"Report generated: {outputPath}");
    }

    private static void GenerateSampleXml(string path, int count)
    {
        var sb = new StringBuilder();
        sb.AppendLine("<Items>");
        for (int i = 1; i <= count; i++)
        {
            sb.AppendLine("  <Item>");
            sb.AppendLine($"    <Name>Item {i}</Name>");
            sb.AppendLine($"    <Value>{i * 10}</Value>");
            sb.AppendLine("  </Item>");
        }
        sb.AppendLine("</Items>");
        File.WriteAllText(path, sb.ToString());
    }

    private static List<Item> LoadItemsWithProgress(string xmlPath, Action<int> progress)
    {
        var items = new List<Item>();

        // First pass to count total items
        int total = 0;
        using (var reader = XmlReader.Create(xmlPath))
        {
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element && reader.Name == "Item")
                    total++;
            }
        }

        if (total == 0)
            return items;

        // Second pass to read items and report progress
        int processed = 0;
        using (var reader = XmlReader.Create(xmlPath))
        {
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element && reader.Name == "Item")
                {
                    var item = new Item();
                    while (reader.Read())
                    {
                        if (reader.NodeType == XmlNodeType.EndElement && reader.Name == "Item")
                            break;

                        if (reader.NodeType == XmlNodeType.Element)
                        {
                            string elementName = reader.Name;
                            if (reader.Read() && reader.NodeType == XmlNodeType.Text)
                            {
                                if (elementName == "Name")
                                    item.Name = reader.Value;
                                else if (elementName == "Value" && int.TryParse(reader.Value, out int val))
                                    item.Value = val;
                            }
                        }
                    }

                    items.Add(item);
                    processed++;
                    int percent = (int)((processed / (double)total) * 100);
                    progress(percent);
                }
            }
        }

        return items;
    }

    private static void CreateTemplate(string path)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Report generated with progress monitoring");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("Name: <<[item.Name]>>, Value: <<[item.Value]>>");
        builder.Writeln("<</foreach>>");
        doc.Save(path);
    }
}

public class ReportModel
{
    public List<Item> Items { get; set; } = new();
}

public class Item
{
    public string Name { get; set; } = string.Empty;
    public int Value { get; set; }
}
