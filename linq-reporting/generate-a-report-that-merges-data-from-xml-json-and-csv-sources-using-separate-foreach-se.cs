using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Item
{
    public int Id { get; set; }
    public string Name { get; set; } = "";
}

public class ReportModel
{
    public List<Item> XmlItems { get; set; } = new();
    public List<Item> JsonItems { get; set; } = new();
    public List<Item> CsvItems { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Ensure code page support for CSV if needed.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        string workDir = Directory.GetCurrentDirectory();
        string dataDir = Path.Combine(workDir, "Data");
        Directory.CreateDirectory(dataDir);
        string outputDir = Path.Combine(workDir, "Output");
        Directory.CreateDirectory(outputDir);

        // Create sample XML file.
        string xmlPath = Path.Combine(dataDir, "sample.xml");
        File.WriteAllText(xmlPath,
@"<Items>
    <Item><Id>1</Id><Name>XML Item One</Name></Item>
    <Item><Id>2</Id><Name>XML Item Two</Name></Item>
</Items>");

        // Create sample JSON file.
        string jsonPath = Path.Combine(dataDir, "sample.json");
        File.WriteAllText(jsonPath,
@"[
    { ""Id"": 101, ""Name"": ""JSON Item Alpha"" },
    { ""Id"": 102, ""Name"": ""JSON Item Beta"" }
]");

        // Create sample CSV file.
        string csvPath = Path.Combine(dataDir, "sample.csv");
        File.WriteAllText(csvPath,
@"Id,Name
201,CSV Item X
202,CSV Item Y");

        // Load data into model.
        ReportModel model = new();

        // XML
        XDocument xDoc = XDocument.Load(xmlPath);
        foreach (XElement elem in xDoc.Root?.Elements("Item") ?? Enumerable.Empty<XElement>())
        {
            int id = (int?)elem.Element("Id") ?? 0;
            string name = (string?)elem.Element("Name") ?? "";
            model.XmlItems.Add(new Item { Id = id, Name = name });
        }

        // JSON
        string jsonContent = File.ReadAllText(jsonPath);
        var jsonItems = JsonConvert.DeserializeObject<List<Item>>(jsonContent);
        if (jsonItems != null)
            model.JsonItems.AddRange(jsonItems);

        // CSV
        var csvLines = File.ReadAllLines(csvPath);
        foreach (var line in csvLines.Skip(1)) // skip header
        {
            var parts = line.Split(',');
            if (parts.Length >= 2 &&
                int.TryParse(parts[0], out int id))
            {
                string name = parts[1];
                model.CsvItems.Add(new Item { Id = id, Name = name });
            }
        }

        // Build template document.
        string templatePath = Path.Combine(workDir, "template.docx");
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);

        builder.Writeln("=== XML Data ===");
        builder.Writeln("<<foreach [item in XmlItems]>>");
        builder.Writeln("- <<[item.Id]>>: <<[item.Name]>>");
        builder.Writeln("<</foreach>>");

        builder.Writeln("\n=== JSON Data ===");
        builder.Writeln("<<foreach [item in JsonItems]>>");
        builder.Writeln("- <<[item.Id]>>: <<[item.Name]>>");
        builder.Writeln("<</foreach>>");

        builder.Writeln("\n=== CSV Data ===");
        builder.Writeln("<<foreach [item in CsvItems]>>");
        builder.Writeln("- <<[item.Id]>>: <<[item.Name]>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // Load template and generate report.
        Document doc = new(templatePath);
        ReportingEngine engine = new();
        engine.BuildReport(doc, model, "model");

        string outputPath = Path.Combine(outputDir, "MergedReport.docx");
        doc.Save(outputPath);
    }
}
