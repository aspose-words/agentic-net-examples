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
        // Register code page provider for CSV handling (if needed).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for the CSV data file, the template, and the output folder.
        string dataPath = "data.csv";
        string templatePath = "template.docx";
        string outputDir = "output";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // 1. Create a sample CSV file with headers: Name,Age,City
        // -----------------------------------------------------------------
        var csvLines = new[]
        {
            "Name,Age,City",
            "Alice,30,New York",
            "Bob,25,London",
            "Charlie,35,Sydney"
        };
        File.WriteAllLines(dataPath, csvLines, Encoding.UTF8);

        // -----------------------------------------------------------------
        // 2. Create a simple Word template that contains LINQ Reporting tags.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // The tags reference the root object name "person".
        builder.Writeln("Report for <<[person.Name]>>");
        builder.Writeln("Age: <<[person.Age]>>");
        builder.Writeln("City: <<[person.City]>>");
        builder.Writeln(); // Add an empty paragraph.

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 3. Load the CSV data into a list of Person objects.
        // -----------------------------------------------------------------
        List<Person> persons = new List<Person>();
        string[] allLines = File.ReadAllLines(dataPath, Encoding.UTF8);
        for (int i = 1; i < allLines.Length; i++) // Skip header line.
        {
            string line = allLines[i];
            if (string.IsNullOrWhiteSpace(line))
                continue;

            string[] parts = line.Split(',');
            if (parts.Length != 3)
                continue;

            persons.Add(new Person
            {
                Name = parts[0].Trim(),
                Age = int.Parse(parts[1].Trim()),
                City = parts[2].Trim()
            });
        }

        // -----------------------------------------------------------------
        // 4. For each person generate an individual document using ReportingEngine.
        // -----------------------------------------------------------------
        foreach (Person person in persons)
        {
            // Load a fresh copy of the template for each iteration.
            Document doc = new Document(templatePath);

            // Build the report. The root object name must match the tag prefix used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, person, "person");

            // Save the generated document with a unique file name.
            string safeName = MakeFileNameSafe(person.Name);
            string outputPath = Path.Combine(outputDir, $"Report_{safeName}.docx");
            doc.Save(outputPath);
        }
    }

    // Helper to replace invalid filename characters.
    private static string MakeFileNameSafe(string name)
    {
        foreach (char c in Path.GetInvalidFileNameChars())
            name = name.Replace(c, '_');
        return name;
    }
}

// Simple data model that matches the template tags.
public class Person
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
    public string City { get; set; } = string.Empty;
}
