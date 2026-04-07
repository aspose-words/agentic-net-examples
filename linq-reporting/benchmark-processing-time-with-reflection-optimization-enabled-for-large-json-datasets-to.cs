using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    // Simple data model for the JSON dataset.
    public class Person
    {
        public int Id { get; set; } = 0;
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; } = 0;
    }

    public static void Main()
    {
        // Register code page provider for any encoding needs.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for temporary files.
        const string jsonPath = "persons.json";
        const string templatePath = "template.docx";
        const string outputPath = "report.docx";

        // 1. Generate a large JSON dataset.
        const int itemCount = 20000; // Adjust size for benchmarking.
        var persons = new List<Person>(itemCount);
        for (int i = 1; i <= itemCount; i++)
        {
            persons.Add(new Person
            {
                Id = i,
                Name = $"Person {i}",
                Age = 20 + (i % 50)
            });
        }

        // Serialize to JSON and write to file.
        string jsonContent = JsonConvert.SerializeObject(persons);
        File.WriteAllText(jsonPath, jsonContent, Encoding.UTF8);

        // 2. Create a LINQ Reporting template programmatically.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Add a simple heading.
        builder.Writeln("Persons Report");
        builder.Writeln("----------------");

        // Insert a foreach tag to iterate over the JSON array.
        builder.Writeln("<<foreach [person in persons]>>");
        builder.Writeln("Id: <<[person.Id]>>, Name: <<[person.Name]>>, Age: <<[person.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // 3. Load the template document.
        Document doc = new Document(templatePath);

        // 4. Prepare the JSON data source.
        JsonDataSource jsonDataSource = new JsonDataSource(jsonPath);

        // 5. Enable reflection optimization.
        ReportingEngine.UseReflectionOptimization = true;

        // 6. Build the report and benchmark the processing time.
        ReportingEngine engine = new ReportingEngine();

        Stopwatch sw = Stopwatch.StartNew();
        engine.BuildReport(doc, jsonDataSource, "persons");
        sw.Stop();

        // 7. Save the generated report.
        doc.Save(outputPath);

        // Output the elapsed time.
        Console.WriteLine($"Report generated in {sw.ElapsedMilliseconds} ms.");
    }
}
