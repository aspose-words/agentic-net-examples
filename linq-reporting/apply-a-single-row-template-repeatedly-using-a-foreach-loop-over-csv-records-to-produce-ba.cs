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
        // Register code page provider for CSV handling.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a simple CSV file.
        string csvPath = Path.Combine(Directory.GetCurrentDirectory(), "data.csv");
        File.WriteAllText(csvPath, "Name,Age,City\r\nAlice,30,New York\r\nBob,25,London\r\nCharlie,35,Sydney");

        // Load CSV records into a list of Person objects.
        List<Person> persons = new();
        using (var reader = new StreamReader(csvPath))
        {
            // Read header.
            string? headerLine = reader.ReadLine();
            if (headerLine == null) return;

            while (!reader.EndOfStream)
            {
                string? line = reader.ReadLine();
                if (string.IsNullOrWhiteSpace(line)) continue;

                string[] parts = line.Split(',');
                if (parts.Length >= 3)
                {
                    persons.Add(new Person
                    {
                        Name = parts[0],
                        Age = int.Parse(parts[1]),
                        City = parts[2]
                    });
                }
            }
        }

        // Create a template document with placeholders for a single Person.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "template.docx");
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);
        builder.Writeln("Report for <<[person.Name]>>");
        builder.Writeln("Age: <<[person.Age]>>");
        builder.Writeln("City: <<[person.City]>>");
        builder.Writeln("-----");
        templateDoc.Save(templatePath);

        // Generate a separate report for each person using the template.
        foreach (Person person in persons)
        {
            Document doc = new(templatePath);
            ReportingEngine engine = new();
            engine.BuildReport(doc, person, "person");

            string outputPath = Path.Combine(outputDir, $"Report_{person.Name}.docx");
            doc.Save(outputPath);
        }
    }
}

public class Person
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
    public string City { get; set; } = "";
}
