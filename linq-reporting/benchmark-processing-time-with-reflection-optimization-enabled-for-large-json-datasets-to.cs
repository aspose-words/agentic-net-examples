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
    public int Id { get; set; } = 0;
    public string Name { get; set; } = "";
    public int Age { get; set; } = 0;
}

public class Program
{
    public static void Main()
    {
        // Register the code page provider required by Aspose.Words for certain encodings.
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Define file names in the current working directory.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        string jsonPath = Path.Combine(Directory.GetCurrentDirectory(), "Data.json");
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");

        // 1. Generate a large JSON dataset.
        const int itemCount = 20000; // Size of the dataset.
        var people = new List<Person>(itemCount);
        for (int i = 1; i <= itemCount; i++)
        {
            people.Add(new Person
            {
                Id = i,
                Name = $"Person {i}",
                Age = 20 + (i % 50)
            });
        }

        // Serialize the dataset to a JSON file.
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(people));

        // 2. Create a LINQ Reporting template programmatically.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("<<foreach [p in persons]>>");
        builder.Writeln("Id: <<[p.Id]>>, Name: <<[p.Name]>>, Age: <<[p.Age]>>");
        builder.Writeln("<</foreach>>");
        templateDoc.Save(templatePath);

        // 3. Load the template document.
        var reportDoc = new Document(templatePath);

        // 4. Prepare the JSON data source.
        var jsonDataSource = new JsonDataSource(jsonPath);

        // 5. Enable reflection optimization for the reporting engine.
        ReportingEngine.UseReflectionOptimization = true;

        // 6. Build the report and measure processing time.
        var engine = new ReportingEngine();

        var stopwatch = Stopwatch.StartNew();
        engine.BuildReport(reportDoc, jsonDataSource, "persons");
        stopwatch.Stop();

        // 7. Save the generated report.
        reportDoc.Save(reportPath);

        // Output the elapsed time.
        Console.WriteLine($"Report generation time with reflection optimization: {stopwatch.ElapsedMilliseconds} ms");
    }
}
