using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV handling.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // -----------------------------------------------------------------
        // 1. Create sample data files (XML, JSON, CSV).
        // -----------------------------------------------------------------
        string xmlPath = "people.xml";
        string jsonPath = "people.json";
        string csvPath = "people.csv";

        File.WriteAllText(xmlPath,
@"<?xml version=""1.0"" encoding=""utf-8""?>
<Persons>
    <Person><Name>John</Name><Age>30</Age></Person>
    <Person><Name>Emma</Name><Age>28</Age></Person>
</Persons>");

        File.WriteAllText(jsonPath,
@"[
    { ""Name"": ""Alice"", ""Age"": 25 },
    { ""Name"": ""Bob"",   ""Age"": 35 }
]");

        File.WriteAllText(csvPath,
@"Name,Age
Mike,40
Sara,22");

        // -----------------------------------------------------------------
        // 2. Build a wrapper model that holds three collections.
        // -----------------------------------------------------------------
        var model = new ReportModel
        {
            Xml = LoadFromXml(xmlPath),
            Json = LoadFromJson(jsonPath),
            Csv = LoadFromCsv(csvPath)
        };

        // -----------------------------------------------------------------
        // 3. Create the template document with three foreach sections.
        // -----------------------------------------------------------------
        string templatePath = "template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("=== XML Data ===");
        builder.Writeln("<<foreach [p in Xml]>>");
        builder.Writeln("<<[p.Name]>> - <<[p.Age]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln();

        builder.Writeln("=== JSON Data ===");
        builder.Writeln("<<foreach [p in Json]>>");
        builder.Writeln("<<[p.Name]>> - <<[p.Age]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln();

        builder.Writeln("=== CSV Data ===");
        builder.Writeln("<<foreach [p in Csv]>>");
        builder.Writeln("<<[p.Name]>> - <<[p.Age]>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 4. Load the template and run the reporting engine.
        // -----------------------------------------------------------------
        var doc = new Document(templatePath);
        var engine = new ReportingEngine();

        // The root object name is "model" – the template tags reference its
        // public properties (Xml, Json, Csv).  Pass the name explicitly.
        engine.BuildReport(doc, model, "model");

        // -----------------------------------------------------------------
        // 5. Save the generated report.
        // -----------------------------------------------------------------
        doc.Save("Report.docx");
    }

    // -----------------------------------------------------------------
    // Helper methods to load data from the three sources.
    // -----------------------------------------------------------------
    private static List<Person> LoadFromXml(string path)
    {
        var list = new List<Person>();
        var xml = new System.Xml.XmlDocument();
        xml.Load(path);
        var nodes = xml.SelectNodes("//Person");
        foreach (System.Xml.XmlNode node in nodes)
        {
            var nameNode = node.SelectSingleNode("Name");
            var ageNode = node.SelectSingleNode("Age");
            list.Add(new Person
            {
                Name = nameNode?.InnerText ?? string.Empty,
                Age = int.TryParse(ageNode?.InnerText, out var a) ? a : 0
            });
        }
        return list;
    }

    private static List<Person> LoadFromJson(string path)
    {
        var json = File.ReadAllText(path);
        return JsonConvert.DeserializeObject<List<Person>>(json) ?? new List<Person>();
    }

    private static List<Person> LoadFromCsv(string path)
    {
        var list = new List<Person>();
        var lines = File.ReadAllLines(path);
        // Assume first line contains headers.
        for (int i = 1; i < lines.Length; i++)
        {
            var parts = lines[i].Split(',');
            if (parts.Length >= 2)
            {
                list.Add(new Person
                {
                    Name = parts[0],
                    Age = int.TryParse(parts[1], out var a) ? a : 0
                });
            }
        }
        return list;
    }
}

// -----------------------------------------------------------------
// Public data model classes – required to be non‑anonymous and
// have public properties that match the template tags.
// -----------------------------------------------------------------
public class ReportModel
{
    public List<Person> Xml { get; set; } = new();
    public List<Person> Json { get; set; } = new();
    public List<Person> Csv { get; set; } = new();
}

public class Person
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
}
