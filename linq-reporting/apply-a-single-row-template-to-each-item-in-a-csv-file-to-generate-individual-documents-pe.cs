using System;
using System.Collections.Generic;
using System.Globalization;
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

        // Prepare sample CSV data.
        string csvPath = "data.csv";
        File.WriteAllText(csvPath, "Name,Age,Email\r\nAlice,30,alice@example.com\r\nBob,25,bob@example.com\r\nCharlie,35,charlie@example.com");

        // Create a simple template document with LINQ Reporting tags.
        string templatePath = "template.docx";
        CreateTemplate(templatePath);

        // Load the template once.
        Document templateDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();

        // Parse CSV into a list of Person objects.
        List<Person> persons = LoadCsv(csvPath);

        // Ensure output directory exists.
        string outputDir = "output";
        Directory.CreateDirectory(outputDir);

        // Generate an individual document for each person.
        for (int i = 0; i < persons.Count; i++)
        {
            Person person = persons[i];
            Document doc = (Document)templateDoc.Clone();

            // Build the report using the person object as the root.
            engine.BuildReport(doc, person, "person");

            string outputPath = Path.Combine(outputDir, $"Person_{i + 1}.docx");
            doc.Save(outputPath);
        }
    }

    // Creates a Word template containing tags that reference the Person model.
    private static void CreateTemplate(string path)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Report for <<[person.Name]>>");
        builder.Writeln("Age: <<[person.Age]>>");
        builder.Writeln("Email: <<[person.Email]>>");

        doc.Save(path);
    }

    // Reads the CSV file and returns a list of Person objects.
    private static List<Person> LoadCsv(string path)
    {
        var persons = new List<Person>();
        using var reader = new StreamReader(path);
        string? headerLine = reader.ReadLine(); // Skip header.
        while (!reader.EndOfStream)
        {
            string? line = reader.ReadLine();
            if (string.IsNullOrWhiteSpace(line))
                continue;

            string[] parts = line.Split(',');
            if (parts.Length < 3)
                continue;

            persons.Add(new Person
            {
                Name = parts[0].Trim(),
                Age = int.Parse(parts[1].Trim(), CultureInfo.InvariantCulture),
                Email = parts[2].Trim()
            });
        }
        return persons;
    }
}

// Public data model matching the template tags.
public class Person
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
    public string Email { get; set; } = string.Empty;
}
