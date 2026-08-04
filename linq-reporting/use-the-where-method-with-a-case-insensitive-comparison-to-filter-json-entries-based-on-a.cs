using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;          // ReportingEngine, JsonDataSource are in this namespace
using Newtonsoft.Json;

public class Person
{
    public string Name { get; set; } = "";
    public string City { get; set; } = "";
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words (required for some environments).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Sample JSON data representing a collection of persons.
        string json = @"
        [
            { ""Name"": ""Alice"",   ""City"": ""London"" },
            { ""Name"": ""Bob"",     ""City"": ""Paris"" },
            { ""Name"": ""Charlie"", ""City"": ""london"" },
            { ""Name"": ""Diana"",   ""City"": ""New York"" }
        ]";

        // Deserialize JSON into a list of Person objects.
        List<Person> allPersons = JsonConvert.DeserializeObject<List<Person>>(json) ?? new List<Person>();

        // Filter persons where City equals "London" (case‑insensitive) using LINQ Where.
        var filteredPersons = allPersons
            .Where(p => string.Equals(p.City, "London", StringComparison.OrdinalIgnoreCase))
            .ToList();

        // Serialize the filtered collection back to JSON.
        string filteredJson = JsonConvert.SerializeObject(filteredPersons);

        // Prepare a JsonDataSource from the filtered JSON string.
        using var jsonStream = new MemoryStream(Encoding.UTF8.GetBytes(filteredJson));
        var jsonDataSource = new JsonDataSource(jsonStream);

        // Create a template document programmatically.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Insert LINQ Reporting tags to iterate over the "persons" data source.
        builder.Writeln("<<foreach [p in persons]>>");
        builder.Writeln("Name: <<[p.Name]>>, City: <<[p.City]>>");
        builder.Writeln("<</foreach>>");

        // Build the report using the ReportingEngine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, jsonDataSource, "persons");

        // Save the generated report.
        doc.Save("FilteredReport.docx");
    }
}
