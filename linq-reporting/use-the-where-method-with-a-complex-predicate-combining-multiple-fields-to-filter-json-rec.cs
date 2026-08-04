using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Prepare working directories.
        string workDir = Directory.GetCurrentDirectory();
        string dataDir = Path.Combine(workDir, "Data");
        Directory.CreateDirectory(dataDir);

        // 1. Create sample JSON data.
        var people = new List<Person>
        {
            new Person { Id = 1, Name = "John Doe", Age = 45, Country = "USA", IsActive = true },
            new Person { Id = 2, Name = "Anna Smith", Age = 28, Country = "UK", IsActive = true },
            new Person { Id = 3, Name = "Mike Johnson", Age = 52, Country = "USA", IsActive = false },
            new Person { Id = 4, Name = "Emily Davis", Age = 33, Country = "USA", IsActive = true },
            new Person { Id = 5, Name = "Li Wei", Age = 40, Country = "CN", IsActive = true }
        };

        string jsonPath = Path.Combine(dataDir, "People.json");
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(people, Formatting.Indented));

        // 2. Build the template document programmatically.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Filtered Persons (Age > 30 && Country == \"USA\" && IsActive):");
        // Loop over all persons; the actual filtering is done with an IF tag.
        builder.Writeln("<<foreach [p in persons]>>");
        // Compare nullable boolean to true to avoid type mismatch.
        builder.Writeln("<<if [p.Age > 30 && p.Country == \"USA\" && p.IsActive == true]>>");
        builder.Writeln("Name: <<[p.Name]>>, Age: <<[p.Age]>>, Country: <<[p.Country]>>");
        builder.Writeln("<</if>>");
        builder.Writeln("<</foreach>>");

        string templatePath = Path.Combine(workDir, "Template.docx");
        templateDoc.Save(templatePath);

        // 3. Load the template for reporting.
        var doc = new Document(templatePath);

        // 4. Create a JSON data source.
        var jsonDataSource = new JsonDataSource(jsonPath);

        // 5. Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(doc, jsonDataSource, "persons");

        // 6. Save the generated report.
        string reportPath = Path.Combine(workDir, "Report.docx");
        doc.Save(reportPath);
    }
}

// Model class for JSON serialization.
public class Person
{
    public int Id { get; set; }
    public string Name { get; set; } = "";
    public int Age { get; set; }
    public string Country { get; set; } = "";
    public bool IsActive { get; set; }
}
