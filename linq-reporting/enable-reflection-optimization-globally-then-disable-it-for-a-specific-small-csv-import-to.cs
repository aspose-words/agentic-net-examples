using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Enable reflection optimization globally.
        ReportingEngine.UseReflectionOptimization = true;

        // Register code page provider for CSV encoding support.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare a small CSV file.
        string csvPath = Path.Combine(Directory.GetCurrentDirectory(), "sample.csv");
        File.WriteAllText(csvPath, "Name,Age\nAlice,30\nBob,25");

        // Load CSV data into a strongly‑typed model.
        var people = LoadCsv(csvPath);
        var model = new ReportModel { Data = people };

        // Create a template document with LINQ Reporting tags.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "template.docx");
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Customer Report");
        builder.Writeln("<<foreach [person in Data]>>");
        builder.Writeln("Name: <<[person.Name]>>, Age: <<[person.Age]>>");
        builder.Writeln("<</foreach>>");
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        var reportDoc = new Document(templatePath);

        // Disable reflection optimization for this small CSV import to reduce overhead.
        ReportingEngine.UseReflectionOptimization = false;

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "report.docx");
        reportDoc.Save(outputPath);
    }

    private static List<Person> LoadCsv(string path)
    {
        var list = new List<Person>();
        foreach (var line in File.ReadAllLines(path))
        {
            // Skip header line.
            if (line.StartsWith("Name", StringComparison.OrdinalIgnoreCase))
                continue;

            var parts = line.Split(',');
            if (parts.Length >= 2)
            {
                list.Add(new Person
                {
                    Name = parts[0],
                    Age = int.TryParse(parts[1], out var age) ? age : 0
                });
            }
        }
        return list;
    }

    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
    }

    public class ReportModel
    {
        public List<Person> Data { get; set; } = new();
    }
}
