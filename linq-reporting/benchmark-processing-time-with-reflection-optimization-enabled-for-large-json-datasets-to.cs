using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    // Simple data model that matches the JSON structure.
    public class Person
    {
        public int Index { get; set; }
        public string Name { get; set; } = "";
        public int Age { get; set; }
    }

    public static void Main()
    {
        // Paths for the temporary files.
        const string jsonPath = "persons.json";
        const string templatePath = "template.docx";
        const string outputPath = "report.docx";

        // 1. Generate a large JSON dataset.
        const int itemCount = 50000;
        var persons = new List<Person>(itemCount);
        for (int i = 0; i < itemCount; i++)
        {
            persons.Add(new Person
            {
                Index = i + 1,
                Name = $"Person_{i + 1}",
                Age = 20 + (i % 50)
            });
        }
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(persons));

        // 2. Create a LINQ Reporting template programmatically.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("<<foreach [person in persons]>>");
        builder.Writeln("<<[person.Index]>> - <<[person.Name]>> - <<[person.Age]>>");
        builder.Writeln("<</foreach>>");
        templateDoc.Save(templatePath);

        // 3. Load the template document.
        var reportDoc = new Document(templatePath);

        // 4. Enable reflection optimization.
        ReportingEngine.UseReflectionOptimization = true;

        // 5. Prepare the JSON data source.
        var jsonDataSource = new JsonDataSource(jsonPath);

        // 6. Build the report and benchmark the processing time.
        var engine = new ReportingEngine();
        var stopwatch = Stopwatch.StartNew();
        engine.BuildReport(reportDoc, jsonDataSource, "persons");
        stopwatch.Stop();

        // 7. Save the generated report.
        reportDoc.Save(outputPath);

        // Output the elapsed time.
        Console.WriteLine($"Report generation time: {stopwatch.ElapsedMilliseconds} ms");
    }
}
