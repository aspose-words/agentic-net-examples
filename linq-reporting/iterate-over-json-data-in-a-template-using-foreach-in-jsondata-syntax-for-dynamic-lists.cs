using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    // Model class matching the JSON objects.
    public class Person
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }
    }

    public static void Main()
    {
        // Enable code page support (required on some platforms).
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // 1. Create sample JSON data file.
        string jsonPath = "people.json";
        var people = new List<Person>
        {
            new Person { Name = "Alice", Age = 30 },
            new Person { Name = "Bob", Age = 25 },
            new Person { Name = "Charlie", Age = 35 }
        };
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(people, Formatting.Indented));

        // 2. Build a template document containing LINQ Reporting tags.
        string templatePath = "Template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Title
        builder.Writeln("People List:");
        // Begin foreach over the JSON data source named 'jsonData'.
        builder.Writeln("<<foreach [person in jsonData]>>");
        // Output each person's details.
        builder.Writeln("- <<[person.Name]>> (Age: <<[person.Age]>>)");
        // End foreach.
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // 3. Load the template for reporting.
        var loadedTemplate = new Document(templatePath);

        // 4. Create a JsonDataSource from the JSON file.
        var jsonDataSource = new JsonDataSource(jsonPath);

        // 5. Build the report using ReportingEngine.
        var engine = new ReportingEngine();
        engine.BuildReport(loadedTemplate, jsonDataSource, "jsonData");

        // 6. Save the generated report.
        string reportPath = "Report.docx";
        loadedTemplate.Save(reportPath);

        // Indicate completion.
        Console.WriteLine($"Report generated: {Path.GetFullPath(reportPath)}");
    }
}
