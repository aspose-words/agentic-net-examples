using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables; // Needed for the Table class
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some environments)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data and write it to a JSON file.
        string jsonPath = "people.json";
        var people = new List<Person>
        {
            new Person { Id = 1, Name = "Alice", Age = 30 },
            new Person { Id = 2, Name = "Bob", Age = 25 },
            new Person { Id = 3, Name = "Charlie", Age = 35 }
        };
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(people, Formatting.Indented));

        // Create a template document programmatically.
        string templatePath = "Template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("People Report");
        builder.Writeln();

        // Begin the foreach loop that will repeat the whole table for each person.
        builder.Writeln("<<foreach [person in persons]>>");

        // Start the table.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Writeln("Id");
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Age");
        builder.EndRow();

        // Data row – tags will be replaced by the reporting engine.
        builder.InsertCell();
        builder.Writeln("<<[person.Id]>>");
        builder.InsertCell();
        builder.Writeln("<<[person.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[person.Age]>>");
        builder.EndRow();

        // Close the table.
        builder.EndTable();

        // End the foreach loop.
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // Load the template document for reporting.
        var reportDoc = new Document(templatePath);

        // Create a JSON data source from the file.
        var jsonDataSource = new JsonDataSource(jsonPath);

        // Build the report using the data source. The root name is "persons".
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, jsonDataSource, "persons");

        // Save the generated report.
        string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}

// Simple data model matching the JSON structure.
public class Person
{
    public int Id { get; set; } = 0;
    public string Name { get; set; } = "";
    public int Age { get; set; } = 0;
}
