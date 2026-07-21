using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare large JSON data
        const int itemCount = 20000;
        var data = new DataWrapper
        {
            Persons = new List<Person>()
        };
        for (int i = 0; i < itemCount; i++)
        {
            data.Persons.Add(new Person
            {
                Name = $"Person_{i:D5}",
                Age = 20 + (i % 50)
            });
        }

        string jsonPath = Path.Combine(Directory.GetCurrentDirectory(), "persons.json");
        File.WriteAllText(jsonPath, JsonConvert.SerializeObject(data));

        // Load data back from JSON (simulating real scenario)
        var jsonContent = File.ReadAllText(jsonPath);
        var model = JsonConvert.DeserializeObject<DataWrapper>(jsonContent)!;

        // Create template document with LINQ Reporting tags
        var template = new Document();
        var builder = new DocumentBuilder(template);

        builder.Writeln("<<foreach [p in data.Persons]>>");
        var table = builder.StartTable();

        // Header row
        builder.InsertCell();
        builder.Writeln("Name");
        builder.InsertCell();
        builder.Writeln("Age");
        builder.EndRow();

        // Data row
        builder.InsertCell();
        builder.Writeln("<<[p.Name]>>");
        builder.InsertCell();
        builder.Writeln("<<[p.Age]>>");
        builder.EndRow();

        builder.EndTable();
        builder.Writeln("<</foreach>>");

        // Enable reflection optimization
        ReportingEngine.UseReflectionOptimization = true;

        var engine = new ReportingEngine();

        // Benchmark the report generation
        var stopwatch = Stopwatch.StartNew();
        bool success = engine.BuildReport(template, model, "data");
        stopwatch.Stop();

        // Save the generated report
        string reportPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        template.Save(reportPath);

        // Output benchmark result
        Console.WriteLine($"Report generation success: {success}");
        Console.WriteLine($"Processing time with reflection optimization: {stopwatch.ElapsedMilliseconds} ms");
        Console.WriteLine($"Report saved to: {reportPath}");
    }
}

public class DataWrapper
{
    public List<Person> Persons { get; set; } = new();
}

public class Person
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
}
