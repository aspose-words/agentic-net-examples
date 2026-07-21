using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class JsonReportExample
{
    public static void Main()
    {
        // Create a working directory for the example files.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "JsonReportExample");
        Directory.CreateDirectory(workDir);

        // 1. Generate a sample JSON file containing a list of Person objects.
        string jsonPath = Path.Combine(workDir, "people.json");
        var people = new List<Person>
        {
            new Person { Name = "Alice", Age = 30, BirthDate = new DateTime(1993, 5, 12) },
            new Person { Name = "Bob",   Age = 45, BirthDate = new DateTime(1978, 11, 3) },
            new Person { Name = "Carol", Age = 27, BirthDate = new DateTime(1996, 2, 20) }
        };
        string jsonContent = System.Text.Json.JsonSerializer.Serialize(
            people,
            new System.Text.Json.JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(jsonPath, jsonContent);

        // 2. Configure JSON loading options to enable loose type inference.
        var jsonLoadOptions = new JsonDataLoadOptions
        {
            SimpleValueParseMode = JsonSimpleValueParseMode.Loose,
            PreserveSpaces = true
        };

        // 3. Create a JsonDataSource from the file using the configured options.
        var jsonDataSource = new JsonDataSource(jsonPath, jsonLoadOptions);

        // 4. Build the template document programmatically.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Insert a foreach loop that iterates over the JSON array named "persons".
        builder.Writeln("<<foreach [person in persons]>>");
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Age: <<[person.Age]>>");
        builder.Writeln("Birth Date: <<[person.BirthDate]>>");
        builder.Writeln("<</foreach>>");

        // 5. Generate the report using ReportingEngine.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;
        engine.BuildReport(doc, jsonDataSource, "persons");

        // 6. Save the resulting document.
        string outputPath = Path.Combine(workDir, "Report.docx");
        doc.Save(outputPath);

        Console.WriteLine($"Report generated at: {outputPath}");
    }

    // Public data model used only for JSON creation; properties are initialized to avoid nullable warnings.
    public class Person
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }
        public DateTime BirthDate { get; set; }
    }
}
