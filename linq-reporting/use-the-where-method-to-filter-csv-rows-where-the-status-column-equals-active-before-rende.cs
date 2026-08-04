using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // 1. Create sample CSV data.
        string csvPath = "data.csv";
        File.WriteAllLines(csvPath, new[]
        {
            "Id,Name,Status",
            "1,John Doe,Active",
            "2,Jane Smith,Inactive",
            "3,Bob Johnson,Active",
            "4,Alice Brown,Inactive"
        });

        // 2. Load CSV into a list of Person objects.
        List<Person> allPersons = new();
        using (var reader = new StreamReader(csvPath))
        {
            // Read header.
            string? headerLine = reader.ReadLine();
            if (headerLine == null) throw new InvalidOperationException("CSV file is empty.");

            // Process each data line.
            while (!reader.EndOfStream)
            {
                string? line = reader.ReadLine();
                if (string.IsNullOrWhiteSpace(line)) continue;

                string[] parts = line.Split(',');
                if (parts.Length != 3) continue; // Skip malformed lines.

                allPersons.Add(new Person
                {
                    Id = int.Parse(parts[0]),
                    Name = parts[1],
                    Status = parts[2]
                });
            }
        }

        // 3. Filter rows where Status == "Active".
        List<Person> activePersons = allPersons
            .Where(p => string.Equals(p.Status, "Active", StringComparison.OrdinalIgnoreCase))
            .ToList();

        // 4. Prepare the data model for the reporting engine.
        ReportModel model = new()
        {
            Persons = activePersons
        };

        // 5. Create the template document programmatically.
        Document template = new();
        DocumentBuilder builder = new(template);

        builder.Writeln("Report of Active Persons:");
        builder.Writeln("<<foreach [person in Persons]>>");
        builder.Writeln("Id: <<[person.Id]>>");
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Status: <<[person.Status]>>");
        builder.Writeln("<</foreach>>");

        // 6. Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new();
        engine.BuildReport(template, model, "model");

        // 7. Save the generated report.
        string outputPath = "ActivePersonsReport.docx";
        template.Save(outputPath);
    }
}

// Data entity representing a row in the CSV.
public class Person
{
    public int Id { get; set; } = 0;
    public string Name { get; set; } = "";
    public string Status { get; set; } = "";
}

// Wrapper model that aligns with the template root object.
public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
}
