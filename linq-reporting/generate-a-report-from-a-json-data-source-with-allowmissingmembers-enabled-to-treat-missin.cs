using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    // Data model used for JSON serialization.
    public class Person
    {
        // Non‑nullable, initialized to avoid warnings.
        public string Name { get; set; } = string.Empty;

        // Nullable to allow missing values.
        public int? Age { get; set; }
    }

    public static void Main()
    {
        // Prepare sample JSON data.
        string jsonPath = "people.json";
        var people = new List<Person>
        {
            new Person { Name = "Alice", Age = 30 },
            new Person { Name = "Bob" } // Age omitted
        };
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(people));

        // Create a template document programmatically.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Report of Persons");
        builder.Writeln("<<foreach [p in persons]>>");
        builder.Writeln("Name: <<[p.Name]>>");
        builder.Writeln("Age: <<[p.Age]>>");
        // 'Address' does not exist in the JSON; it will be treated as null.
        builder.Writeln("Address: <<[p.Address]>>");
        builder.Writeln("<</foreach>>");

        // Configure the reporting engine to allow missing members.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.AllowMissingMembers;
        engine.MissingMemberMessage = "N/A";

        // Load JSON data as a data source.
        JsonDataSource jsonDataSource = new JsonDataSource(jsonPath);

        // Build the report. The data source name ("persons") must match the tag used in the template.
        engine.BuildReport(doc, jsonDataSource, "persons");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}
