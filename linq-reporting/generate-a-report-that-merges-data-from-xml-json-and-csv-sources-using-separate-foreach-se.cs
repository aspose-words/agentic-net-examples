using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV handling.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create sample data files.
        CreateSampleXml("data.xml");
        CreateSampleJson("data.json");
        CreateSampleCsv("data.csv");

        // Load data from the files into strongly typed collections.
        var xmlItems = LoadXmlItems("data.xml");
        var jsonItems = LoadJsonItems("data.json");
        var csvItems = LoadCsvItems("data.csv");

        // Assemble the root model that will be passed to the reporting engine.
        var model = new ReportModel
        {
            XmlItems = xmlItems,
            JsonItems = jsonItems,
            CsvItems = csvItems
        };

        // Build the LINQ Reporting template programmatically.
        var templatePath = "template.docx";
        BuildTemplate(templatePath);

        // Load the template document.
        var doc = new Document(templatePath);

        // Generate the report.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the final report.
        var outputPath = "report.docx";
        doc.Save(outputPath);
    }

    // Creates a simple XML file with a list of items.
    private static void CreateSampleXml(string path)
    {
        var xml = new XDocument(
            new XElement("Items",
                new XElement("Item",
                    new XElement("Id", 1),
                    new XElement("Name", "Alpha")),
                new XElement("Item",
                    new XElement("Id", 2),
                    new XElement("Name", "Beta")),
                new XElement("Item",
                    new XElement("Id", 3),
                    new XElement("Name", "Gamma"))
            )
        );
        xml.Save(path);
    }

    // Creates a simple JSON file with a list of objects.
    private static void CreateSampleJson(string path)
    {
        var items = new List<JsonItem>
        {
            new JsonItem { Id = 101, Description = "First item" },
            new JsonItem { Id = 102, Description = "Second item" },
            new JsonItem { Id = 103, Description = "Third item" }
        };
        var json = JsonConvert.SerializeObject(items, Formatting.Indented);
        File.WriteAllText(path, json);
    }

    // Creates a simple CSV file with a header and a few rows.
    private static void CreateSampleCsv(string path)
    {
        var lines = new[]
        {
            "Id,Value",
            "201,Apple",
            "202,Banana",
            "203,Cherry"
        };
        File.WriteAllLines(path, lines, Encoding.UTF8);
    }

    // Loads XML items into a list of XmlItem objects.
    private static List<XmlItem> LoadXmlItems(string path)
    {
        var doc = XDocument.Load(path);
        return doc.Root!
                  .Elements("Item")
                  .Select(x => new XmlItem
                  {
                      Id = (int)x.Element("Id")!,
                      Name = (string)x.Element("Name")!
                  })
                  .ToList();
    }

    // Loads JSON items into a list of JsonItem objects.
    private static List<JsonItem> LoadJsonItems(string path)
    {
        var json = File.ReadAllText(path);
        return JsonConvert.DeserializeObject<List<JsonItem>>(json) ?? new List<JsonItem>();
    }

    // Loads CSV items into a list of CsvItem objects.
    private static List<CsvItem> LoadCsvItems(string path)
    {
        var lines = File.ReadAllLines(path);
        var items = new List<CsvItem>();
        for (int i = 1; i < lines.Length; i++) // Skip header
        {
            var parts = lines[i].Split(',');
            if (parts.Length >= 2 &&
                int.TryParse(parts[0], out int id))
            {
                items.Add(new CsvItem
                {
                    Id = id,
                    Value = parts[1]
                });
            }
        }
        return items;
    }

    // Builds the LINQ Reporting template with three separate foreach sections.
    private static void BuildTemplate(string path)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // XML data section.
        builder.Writeln("XML Data:");
        builder.Writeln("<<foreach [item in XmlItems]>>");
        builder.Writeln("- Id: <<[item.Id]>>, Name: <<[item.Name]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln();

        // JSON data section.
        builder.Writeln("JSON Data:");
        builder.Writeln("<<foreach [item in JsonItems]>>");
        builder.Writeln("- Id: <<[item.Id]>>, Description: <<[item.Description]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln();

        // CSV data section.
        builder.Writeln("CSV Data:");
        builder.Writeln("<<foreach [item in CsvItems]>>");
        builder.Writeln("- Id: <<[item.Id]>>, Value: <<[item.Value]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(path);
    }
}

// Root model containing collections from each source.
public class ReportModel
{
    public List<XmlItem> XmlItems { get; set; } = new();
    public List<JsonItem> JsonItems { get; set; } = new();
    public List<CsvItem> CsvItems { get; set; } = new();
}

// Model for XML items.
public class XmlItem
{
    public int Id { get; set; }
    public string Name { get; set; } = string.Empty;
}

// Model for JSON items.
public class JsonItem
{
    public int Id { get; set; }
    public string Description { get; set; } = string.Empty;
}

// Model for CSV items.
public class CsvItem
{
    public int Id { get; set; }
    public string Value { get; set; } = string.Empty;
}
