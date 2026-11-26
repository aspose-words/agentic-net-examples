using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Person
{
    public string Name { get; set; } = "John Doe";
    public int Age { get; set; } = 30;
}

public class ReportModel
{
    public Person Person { get; set; } = new Person();

    public List<Person> People { get; set; } = new()
    {
        new Person { Name = "Alice", Age = 28 },
        new Person { Name = "Bob", Age = 35 }
    };

    // Serialized JSON representation of the People collection.
    public string PeopleJson => JsonConvert.SerializeObject(People);
}

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Create a template document.
        string templatePath = Path.Combine(outputDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("=== LINQ Reporting Example ===");
        builder.Writeln("Person Name: <<[model.Person.Name]>>");
        builder.Writeln("Person Age: <<[model.Person.Age]>>");
        builder.Writeln();
        builder.Writeln("All People (JSON serialization using JsonConvert):");
        builder.Writeln("<<[model.PeopleJson]>>");

        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document doc = new Document(templatePath);

        // Create the data model.
        ReportModel model = new ReportModel();

        // Configure the reporting engine.
        ReportingEngine engine = new ReportingEngine();

        // Build the report.
        bool success = engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string resultPath = Path.Combine(outputDir, "result.docx");
        doc.Save(resultPath);

        // Output result information.
        Console.WriteLine($"Report generation successful: {success}");
        Console.WriteLine($"Template saved to: {templatePath}");
        Console.WriteLine($"Result saved to: {resultPath}");
    }
}
