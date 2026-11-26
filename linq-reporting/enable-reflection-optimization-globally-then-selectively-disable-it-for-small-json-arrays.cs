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
        // Ensure the output directory exists.
        Directory.CreateDirectory("Output");

        // 1. Create a LINQ Reporting template.
        string templatePath = "Template.docx";
        CreateTemplate(templatePath);

        // 2. Prepare sample JSON data files.
        string largeJsonPath = "large.json";
        string smallJsonPath = "small.json";
        CreateSampleJson(largeJsonPath, 100); // 100 items – considered large.
        CreateSampleJson(smallJsonPath, 3);   // 3 items – considered small.

        // 3. Deserialize JSON into model objects.
        ReportModel largeModel = JsonConvert.DeserializeObject<ReportModel>(File.ReadAllText(largeJsonPath))!;
        ReportModel smallModel = JsonConvert.DeserializeObject<ReportModel>(File.ReadAllText(smallJsonPath))!;

        // 4. Enable reflection optimization globally.
        ReportingEngine.UseReflectionOptimization = true;

        // 5. Build report for large data (optimization stays enabled).
        Document largeDoc = new Document(templatePath);
        ReportingEngine largeEngine = new ReportingEngine();
        largeEngine.BuildReport(largeDoc, largeModel, "model");
        largeDoc.Save("Output/LargeReport.docx");

        // 6. Disable reflection optimization for small data.
        ReportingEngine.UseReflectionOptimization = false;
        Document smallDoc = new Document(templatePath);
        ReportingEngine smallEngine = new ReportingEngine();
        smallEngine.BuildReport(smallDoc, smallModel, "model");
        smallDoc.Save("Output/SmallReport.docx");

        // 7. Re‑enable optimization if further processing is needed.
        ReportingEngine.UseReflectionOptimization = true;
    }

    // Creates a simple Word template with LINQ Reporting tags.
    private static void CreateTemplate(string path)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Person List:");
        builder.Writeln("<<foreach [p in Persons]>>");
        builder.Writeln("Name: <<[p.Name]>>, Age: <<[p.Age]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(path);
    }

    // Generates a JSON file containing a list of persons.
    private static void CreateSampleJson(string path, int count)
    {
        var model = new ReportModel
        {
            Persons = new List<Person>()
        };

        for (int i = 1; i <= count; i++)
        {
            model.Persons.Add(new Person
            {
                Name = $"Person {i}",
                Age = 20 + i % 30
            });
        }

        string json = JsonConvert.SerializeObject(model, Formatting.Indented);
        File.WriteAllText(path, json);
    }
}

// Root data model for the report.
public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
}

// Simple person class referenced by the template.
public class Person
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
}
