using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Person
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
}

public class Program
{
    public static void Main()
    {
        // Prepare sample JSON data.
        string jsonPath = "people.json";
        string jsonContent = @"[
            { ""Name"": ""Alice"", ""Age"": 30 },
            { ""Name"": ""bob"",   ""Age"": 25 },
            { ""Name"": ""ALICE"", ""Age"": 28 }
        ]";
        File.WriteAllText(jsonPath, jsonContent);

        // Deserialize JSON into a list of Person objects.
        List<Person> allPersons = JsonConvert.DeserializeObject<List<Person>>(File.ReadAllText(jsonPath)) ?? new List<Person>();

        // Filter entries where Name equals "alice" (case‑insensitive).
        List<Person> filteredPersons = allPersons
            .Where(p => string.Equals(p.Name, "alice", StringComparison.OrdinalIgnoreCase))
            .ToList();

        // Create a template document programmatically.
        string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Filtered Persons:");
        builder.Writeln("<<foreach [p in persons]>>");
        builder.Writeln("Name: <<[p.Name]>>, Age: <<[p.Age]>>");
        builder.Writeln("<</foreach>>");
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // Build the report using the filtered list as the data source.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(reportDoc, filteredPersons, "persons");

        // Save the final report.
        string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}
