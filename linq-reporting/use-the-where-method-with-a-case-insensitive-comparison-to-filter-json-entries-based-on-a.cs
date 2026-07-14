using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Prepare sample JSON data.
        string jsonPath = "people.json";
        File.WriteAllText(jsonPath,
            @"[
                { ""Name"": ""Alice"",   ""Category"": ""Admin"" },
                { ""Name"": ""Bob"",     ""Category"": ""User"" },
                { ""Name"": ""Charlie"", ""Category"": ""admin"" },
                { ""Name"": ""Diana"",   ""Category"": ""Guest"" }
            ]");

        // Deserialize JSON into a list of Person objects.
        List<Person> allPeople = JsonConvert.DeserializeObject<List<Person>>(File.ReadAllText(jsonPath))!;

        // Filter entries where Category equals \"admin\" (case‑insensitive).
        List<Person> filtered = allPeople
            .Where(p => string.Equals(p.Category, "admin", StringComparison.OrdinalIgnoreCase))
            .ToList();

        // Wrap the filtered collection for the reporting engine.
        ReportModel model = new()
        {
            People = filtered
        };

        // Create a Word template with LINQ Reporting tags.
        Document doc = new();
        DocumentBuilder builder = new(doc);

        builder.Writeln("<<foreach [person in People]>>");
        builder.Writeln("Name: <<[person.Name]>>, Category: <<[person.Category]>>");
        builder.Writeln("<</foreach>>");

        // Build the report.
        ReportingEngine engine = new();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}

// Data entity representing a person.
public class Person
{
    public string Name { get; set; } = "";
    public string Category { get; set; } = "";
}

// Wrapper class used as the root data source for the report.
public class ReportModel
{
    public List<Person> People { get; set; } = new();
}
