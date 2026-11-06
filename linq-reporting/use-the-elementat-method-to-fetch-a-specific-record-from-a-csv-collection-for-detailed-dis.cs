using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare output directory
        string outputDir = "output";
        Directory.CreateDirectory(outputDir);

        // Create a sample CSV file
        string csvPath = Path.Combine(outputDir, "people.csv");
        File.WriteAllText(csvPath,
            "Id,Name,Age\n" +
            "1,John Doe,30\n" +
            "2,Jane Smith,25\n" +
            "3,Bob Johnson,40");

        // Load CSV data into a strongly‑typed list
        List<Person> persons = File.ReadAllLines(csvPath)
            .Skip(1) // Skip header
            .Select(line =>
            {
                var parts = line.Split(',');
                return new Person
                {
                    Id = int.Parse(parts[0]),
                    Name = parts[1],
                    Age = int.Parse(parts[2])
                };
            })
            .ToList();

        // Fetch the third record (index 2) using ElementAt
        Person specificRecord = persons.ElementAt(2);

        // Create a template document programmatically
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Detailed Record:");
        builder.Writeln("Id: <<[record.Id]>>");
        builder.Writeln("Name: <<[record.Name]>>");
        builder.Writeln("Age: <<[record.Age]>>");

        // Build the report using the specific record as the root object
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, specificRecord, "record");

        // Save the generated report
        string outPath = Path.Combine(outputDir, "Report.docx");
        doc.Save(outPath);
    }
}

// Public data model for CSV rows
public class Person
{
    public int Id { get; set; } = 0;
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; } = 0;
}
