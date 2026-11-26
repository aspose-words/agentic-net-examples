using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Person
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
    public string City { get; set; } = "";
    public bool IsActive { get; set; }
}

public class ReportModel
{
    public List<Person> FilteredPersons { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample JSON data.
        string jsonPath = "data.json";
        var sampleData = new List<Person>
        {
            new() { Name = "Alice", Age = 28, City = "New York", IsActive = true },
            new() { Name = "Bob", Age = 35, City = "New York", IsActive = false },
            new() { Name = "Charlie", Age = 42, City = "Los Angeles", IsActive = true },
            new() { Name = "Diana", Age = 31, City = "New York", IsActive = true }
        };
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(sampleData, Formatting.Indented));

        // Load JSON into objects.
        var allPersons = JsonConvert.DeserializeObject<List<Person>>(File.ReadAllText(jsonPath)) ?? new();

        // Apply complex predicate: Age > 30, City == "New York", and IsActive == true.
        var filtered = allPersons
            .Where(p => p.Age > 30 && p.City == "New York" && p.IsActive)
            .ToList();

        // Prepare the model for the report.
        var model = new ReportModel
        {
            FilteredPersons = filtered
        };

        // Create the LINQ Reporting template programmatically.
        string templatePath = "template.docx";
        var builder = new DocumentBuilder();
        builder.Writeln("Filtered Persons:");
        builder.Writeln("<<foreach [p in FilteredPersons]>>");
        builder.Writeln("Name: <<[p.Name]>>, Age: <<[p.Age]>>, City: <<[p.City]>>");
        builder.Writeln("<</foreach>>");
        var templateDoc = builder.Document;
        templateDoc.Save(templatePath);

        // Load the template and build the report.
        var doc = new Document(templatePath);
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string outputPath = "output.docx";
        doc.Save(outputPath);
    }
}
