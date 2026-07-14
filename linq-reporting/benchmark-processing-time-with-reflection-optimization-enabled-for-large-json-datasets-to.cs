using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Person
{
    public int Id { get; set; }
    public string Name { get; set; } = "";
    public int Age { get; set; }
}

public class DataModel
{
    public List<Person> Persons { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words if needed.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare directories.
        string workDir = Directory.GetCurrentDirectory();
        string dataFile = Path.Combine(workDir, "persons.json");
        string templateFile = Path.Combine(workDir, "template.docx");
        string resultFile = Path.Combine(workDir, "result.docx");

        // Generate a large JSON dataset.
        const int itemCount = 20000; // Adjust for desired size.
        var model = new DataModel();
        for (int i = 1; i <= itemCount; i++)
        {
            model.Persons.Add(new Person
            {
                Id = i,
                Name = $"Person {i}",
                Age = 20 + (i % 50)
            });
        }
        // Serialize to JSON file.
        File.WriteAllText(dataFile, JsonConvert.SerializeObject(model));

        // Create a template document with LINQ Reporting tags.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("<<foreach [person in Persons]>>");
        builder.Writeln("Id: <<[person.Id]>>");
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Age: <<[person.Age]>>");
        builder.Writeln("<</foreach>>");
        doc.Save(templateFile);

        // Load the template.
        var template = new Document(templateFile);

        // Load JSON data source.
        var jsonDataSource = new JsonDataSource(dataFile);

        // Enable reflection optimization.
        ReportingEngine.UseReflectionOptimization = true;

        var engine = new ReportingEngine();

        // Benchmark the BuildReport call.
        var stopwatch = Stopwatch.StartNew();
        engine.BuildReport(template, jsonDataSource, "Persons");
        stopwatch.Stop();

        // Save the generated report.
        template.Save(resultFile);

        // Output the elapsed time.
        Console.WriteLine($"Report generation time: {stopwatch.ElapsedMilliseconds} ms");
    }
}
