using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;
using System.Text;

public class Person
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
    public string Email { get; set; } = "";
}

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for Aspose.Words on .NET Core).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data.
        var people = new List<Person>
        {
            new Person { Name = "John Doe", Age = 30, Email = "john.doe@example.com" },
            new Person { Name = "Jane Smith", Age = 25, Email = "jane.smith@example.com" },
            new Person { Name = "Bob Johnson", Age = 40, Email = "bob.johnson@example.com" }
        };

        // Serialize data to a JSON file.
        const string jsonPath = "people.json";
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(people, Formatting.Indented));

        // Create a template document programmatically.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // LINQ Reporting tags: iterate over each JSON object and output its details.
        builder.Writeln("<<foreach [person in persons]>>");
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Age: <<[person.Age]>>");
        builder.Writeln("Email: <<[person.Email]>>");
        builder.Writeln("<</foreach>>");

        // Load JSON data as a data source.
        var jsonDataSource = new JsonDataSource(jsonPath);

        // Build the report.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(doc, jsonDataSource, "persons");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}
