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
        // Register code page provider (required for some environments).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data.
        var people = new List<Person>
        {
            new Person { Name = "Alice", Age = 30, IsMember = true, JoinDate = new DateTime(2020, 5, 12) },
            new Person { Name = "Bob", Age = 45, IsMember = false, JoinDate = new DateTime(2018, 11, 3) },
            new Person { Name = "Charlie", Age = 28, IsMember = true, JoinDate = new DateTime(2021, 2, 20) }
        };

        // Serialize data to JSON and write to a file.
        string jsonPath = "people.json";
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(people, Formatting.Indented));

        // Create a template document with LINQ Reporting tags.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("<<foreach [person in persons]>>");
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Age: <<[person.Age]>>");
        builder.Writeln("Member: <<[person.IsMember]>>");
        builder.Writeln("Joined: <<[person.JoinDate]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        string templatePath = "template.docx";
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        var reportDoc = new Document(templatePath);

        // Enable type inference via JsonDataLoadOptions (default behavior, but we set it explicitly).
        var jsonOptions = new JsonDataLoadOptions
        {
            // Loose parsing allows the engine to infer types such as int, bool, DateTime, etc.
            SimpleValueParseMode = JsonSimpleValueParseMode.Loose,
            PreserveSpaces = true,
            AlwaysGenerateRootObject = true
        };

        // Create the JSON data source.
        var jsonDataSource = new JsonDataSource(jsonPath, jsonOptions);

        // Build the report.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // default options
        engine.BuildReport(reportDoc, jsonDataSource, "persons");

        // Save the generated report.
        string outputPath = "report.docx";
        reportDoc.Save(outputPath);
    }
}

// Data model class with public properties (required for LINQ Reporting).
public class Person
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
    public bool IsMember { get; set; }
    public DateTime JoinDate { get; set; }
}
