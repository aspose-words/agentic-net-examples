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
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare folders.
        string outputDir = "output";
        Directory.CreateDirectory(outputDir);

        // 1. Create the LINQ Reporting template.
        string templatePath = Path.Combine(outputDir, "template.docx");
        CreateTemplate(templatePath);

        // 2. Generate a large JSON dataset.
        string jsonPath = Path.Combine(outputDir, "data.json");
        GenerateLargeJson(jsonPath, 20000); // 20,000 records

        // 3. Load the template document.
        Document doc = new Document(templatePath);

        // 4. Load JSON data source.
        JsonDataSource jsonDataSource = new JsonDataSource(jsonPath);

        // 5. Enable reflection optimization.
        ReportingEngine.UseReflectionOptimization = true;

        // 6. Build the report and benchmark the time.
        ReportingEngine engine = new ReportingEngine();

        Stopwatch sw = Stopwatch.StartNew();
        engine.BuildReport(doc, jsonDataSource, "data");
        sw.Stop();

        // 7. Save the generated report.
        string reportPath = Path.Combine(outputDir, "report.docx");
        doc.Save(reportPath);

        // 8. Output the elapsed time.
        Console.WriteLine($"Report generated in {sw.ElapsedMilliseconds} ms with reflection optimization enabled.");
    }

    private static void CreateTemplate(string path)
    {
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Write LINQ Reporting tags.
        builder.Writeln("<<foreach [item in data]>>");
        builder.Writeln("<<[item.Name]>> - <<[item.Age]>>");
        builder.Writeln("<</foreach>>");

        // Save the template.
        templateDoc.Save(path);
    }

    private static void GenerateLargeJson(string path, int count)
    {
        List<Person> persons = new List<Person>(count);
        for (int i = 0; i < count; i++)
        {
            persons.Add(new Person
            {
                Name = $"Person_{i + 1}",
                Age = 20 + (i % 50)
            });
        }

        string json = JsonConvert.SerializeObject(persons);
        File.WriteAllText(path, json, Encoding.UTF8);
    }
}

// Public data model for JSON serialization.
public class Person
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
}
