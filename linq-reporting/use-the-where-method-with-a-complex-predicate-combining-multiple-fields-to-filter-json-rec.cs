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
        // Ensure the working directory exists.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(workDir);

        // 1. Create sample JSON data.
        string jsonPath = Path.Combine(workDir, "people.json");
        var sampleData = new List<Person>
        {
            new Person { Name = "John Doe", Age = 45, Country = "USA", IsActive = true },
            new Person { Name = "Jane Smith", Age = 28, Country = "USA", IsActive = true },
            new Person { Name = "Carlos Ruiz", Age = 52, Country = "Spain", IsActive = true },
            new Person { Name = "Emily Zhang", Age = 37, Country = "USA", IsActive = false },
            new Person { Name = "Anna Müller", Age = 33, Country = "Germany", IsActive = true }
        };
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(sampleData, Formatting.Indented));

        // 2. Load JSON and apply a complex Where predicate.
        List<Person> allPersons = JsonConvert.DeserializeObject<List<Person>>(File.ReadAllText(jsonPath))!;
        List<Person> filteredPersons = allPersons
            .Where(p => p.Age > 30 && p.Country == "USA" && p.IsActive)
            .ToList();

        // 3. Create a wrapper model for the reporting engine.
        var model = new ReportModel { Persons = filteredPersons };

        // 4. Build the LINQ Reporting template programmatically.
        string templatePath = Path.Combine(workDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Filtered Persons (Age > 30, Country = USA, IsActive = true):");
        builder.Writeln("<<foreach [p in Persons]>>");
        builder.Writeln("Name: <<[p.Name]>>");
        builder.Writeln("Age: <<[p.Age]>>");
        builder.Writeln("Country: <<[p.Country]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // 5. Load the template and generate the report.
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(reportDoc, model, "model");

        // 6. Save the final report.
        string outputPath = Path.Combine(workDir, "FilteredReport.docx");
        reportDoc.Save(outputPath);
    }
}

// Data entity representing a person.
public class Person
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
    public string Country { get; set; } = string.Empty;
    public bool IsActive { get; set; }
}

// Wrapper model exposing the filtered collection to the template.
public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
}
