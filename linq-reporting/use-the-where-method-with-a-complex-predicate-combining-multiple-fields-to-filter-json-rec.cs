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
        // Ensure code pages are available (required by Aspose.Words for some encodings).
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // 1. Create sample JSON data file.
        string jsonPath = "people.json";
        var samplePeople = new List<Person>
        {
            new Person { Name = "Alice", Age = 28, City = "New York", IsActive = true },
            new Person { Name = "Bob",   Age = 45, City = "Chicago",   IsActive = false },
            new Person { Name = "Carol", Age = 34, City = "New York", IsActive = true },
            new Person { Name = "Dave",  Age = 52, City = "Los Angeles", IsActive = true },
            new Person { Name = "Eve",   Age = 23, City = "New York", IsActive = false }
        };
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(samplePeople, Formatting.Indented));

        // 2. Load JSON data into objects.
        var allPeople = JsonConvert.DeserializeObject<List<Person>>(File.ReadAllText(jsonPath)) ?? new List<Person>();

        // 3. Apply a complex LINQ Where predicate to filter records dynamically.
        //    Keep only active persons older than 30 who live in New York.
        var filteredPeople = allPeople
            .Where(p => p.IsActive && p.Age > 30 && string.Equals(p.City, "New York", StringComparison.OrdinalIgnoreCase))
            .ToList();

        // 4. Prepare the reporting model.
        var model = new ReportModel { Persons = filteredPeople };

        // 5. Create a template document with LINQ Reporting tags.
        string templatePath = "template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Filtered Persons Report");
        builder.Writeln("<<foreach [person in Persons]>>");
        builder.Writeln("- <<[person.Name]>> (Age: <<[person.Age]>>, City: <<[person.City]>>)");
        builder.Writeln("<</foreach>>");
        templateDoc.Save(templatePath);

        // 6. Load the template document for reporting.
        var loadedTemplate = new Document(templatePath);

        // 7. Build the report using ReportingEngine.
        var engine = new ReportingEngine();
        engine.BuildReport(loadedTemplate, model, "model");

        // 8. Save the generated report.
        string outputPath = "FilteredReport.docx";
        loadedTemplate.Save(outputPath);
    }
}

// Public data model representing a person.
public class Person
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
    public string City { get; set; } = string.Empty;
    public bool IsActive { get; set; }
}

// Wrapper class used as the root data source for the report.
public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
}
