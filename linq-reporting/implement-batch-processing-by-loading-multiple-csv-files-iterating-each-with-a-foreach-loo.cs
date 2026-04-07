using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    public int Id { get; set; }
    public string Name { get; set; } = string.Empty;
}

public class Program
{
    // Loads a simple CSV file (with header line) into a list of Person objects.
    private static List<Person> LoadCsv(string csvPath)
    {
        var persons = new List<Person>();
        var lines = File.ReadAllLines(csvPath);
        // Skip header.
        for (int i = 1; i < lines.Length; i++)
        {
            var parts = lines[i].Split(',');
            if (parts.Length >= 2 &&
                int.TryParse(parts[0], out int id))
            {
                persons.Add(new Person { Id = id, Name = parts[1] });
            }
        }
        return persons;
    }

    public static void Main()
    {
        // Ensure the working folder exists.
        string dataFolder = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        Directory.CreateDirectory(dataFolder);

        // Create two sample CSV files with headers.
        string csvPath1 = Path.Combine(dataFolder, "people1.csv");
        string csvPath2 = Path.Combine(dataFolder, "people2.csv");

        File.WriteAllText(csvPath1,
            "Id,Name\n" +
            "1,John Doe\n" +
            "2,Jane Smith\n");

        File.WriteAllText(csvPath2,
            "Id,Name\n" +
            "3,Albert Einstein\n" +
            "4,Marie Curie\n");

        // -----------------------------------------------------------------
        // Step 1. Create a template document programmatically.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        builder.Writeln("Batch CSV Report");
        builder.Writeln();

        // First CSV file loop.
        builder.Writeln("<<foreach [p in people1]>>");
        builder.Writeln("File 1 - Id: <<[p.Id]>>, Name: <<[p.Name]>>");
        builder.Writeln("<</foreach>>");
        builder.Writeln();

        // Second CSV file loop.
        builder.Writeln("<<foreach [p in people2]>>");
        builder.Writeln("File 2 - Id: <<[p.Id]>>, Name: <<[p.Name]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk (required before building the report).
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // Step 2. Load the template and the CSV data sources.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);

        // Load each CSV file into a strongly‑typed collection.
        List<Person> people1 = LoadCsv(csvPath1);
        List<Person> people2 = LoadCsv(csvPath2);

        // -----------------------------------------------------------------
        // Step 3. Build the report using multiple data sources.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();

        // The data source names must match the names used in the template tags.
        engine.BuildReport(doc,
            new object[] { people1, people2 },
            new[] { "people1", "people2" });

        // -----------------------------------------------------------------
        // Step 4. Save the generated report.
        // -----------------------------------------------------------------
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "BatchReport.docx");
        doc.Save(outputPath);

        Console.WriteLine($"Report generated: {outputPath}");
    }
}
