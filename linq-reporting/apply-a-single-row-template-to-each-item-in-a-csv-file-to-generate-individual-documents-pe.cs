using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    // Simple data model that matches the CSV columns.
    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
        public string City { get; set; } = string.Empty;
    }

    public static void Main()
    {
        // Register code page provider for possible non‑UTF8 CSV files.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample CSV data.
        string csvPath = "people.csv";
        File.WriteAllLines(csvPath, new[]
        {
            "Name,Age,City",
            "Alice,30,New York",
            "Bob,25,London",
            "Charlie,35,Sydney"
        });

        // Load CSV into a list of Person objects.
        List<Person> people = LoadCsv(csvPath);

        // Create a single‑row template programmatically.
        string templatePath = "template.docx";
        CreateTemplate(templatePath);

        // Generate an individual document for each record.
        for (int i = 0; i < people.Count; i++)
        {
            // Load a fresh copy of the template for each iteration.
            Document doc = new Document(templatePath);

            // Build the report using the current Person as the data source.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, people[i], "model");

            // Save the generated document with a distinct file name.
            string outputPath = $"Person_{i + 1}_{people[i].Name}.docx";
            doc.Save(outputPath);
        }
    }

    // Reads a CSV file with a header row and returns a list of Person objects.
    private static List<Person> LoadCsv(string path)
    {
        var result = new List<Person>();
        string[] lines = File.ReadAllLines(path);
        if (lines.Length < 2) return result; // No data.

        // Assume first line contains headers.
        for (int i = 1; i < lines.Length; i++)
        {
            string line = lines[i];
            if (string.IsNullOrWhiteSpace(line)) continue;

            string[] parts = line.Split(',');
            if (parts.Length != 3) continue; // Skip malformed rows.

            var person = new Person
            {
                Name = parts[0].Trim(),
                Age = int.TryParse(parts[1].Trim(), out int age) ? age : 0,
                City = parts[2].Trim()
            };
            result.Add(person);
        }
        return result;
    }

    // Creates a simple Word template containing LINQ Reporting tags.
    private static void CreateTemplate(string path)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Personal Report");
        builder.Writeln("----------------");
        builder.Writeln("Name: <<[model.Name]>>");
        builder.Writeln("Age:  <<[model.Age]>>");
        builder.Writeln("City: <<[model.City]>>");

        doc.Save(path);
    }
}
