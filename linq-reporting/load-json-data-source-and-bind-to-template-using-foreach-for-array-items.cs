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
        // Register code page provider for Aspose.Words (required on .NET Core).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data.
        var persons = new List<Person>
        {
            new Person { Name = "Alice", Age = 30 },
            new Person { Name = "Bob", Age = 25 },
            new Person { Name = "Charlie", Age = 28 }
        };

        // Serialize the list directly to JSON (root is an array).
        string jsonPath = Path.Combine(Directory.GetCurrentDirectory(), "data.json");
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(persons, Formatting.Indented));

        // Create a template document with LINQ Reporting tags.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Person List:");
        builder.Writeln();

        // Iterate over the root array named "data".
        builder.Writeln("<<foreach [p in data]>>");
        builder.Writeln("<<[p.Name]>> - <<[p.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        var reportDoc = new Document(templatePath);

        // Create a JsonDataSource pointing to the JSON file.
        var jsonDataSource = new JsonDataSource(jsonPath);

        // Build the report using the JSON data source.
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, jsonDataSource, "data");

        // Save the generated report.
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        reportDoc.Save(reportPath);
    }
}

// Simple data model used for JSON creation.
public class Person
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
}
