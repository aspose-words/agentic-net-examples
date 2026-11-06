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
        // Register code page provider for possible CSV encoding handling.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample CSV data.
        string csvPath = "data.csv";
        File.WriteAllText(csvPath,
            "Name,Age,City\n" +
            "Alice,30,New York\n" +
            "Bob,25,Los Angeles\n" +
            "Charlie,35,Chicago");

        // Create a simple template with LINQ Reporting tags.
        string templatePath = "template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Report for <<[record.Name]>>");
        builder.Writeln("Age: <<[record.Age]>>");
        builder.Writeln("City: <<[record.City]>>");
        templateDoc.Save(templatePath);

        // Load CSV records into a list of Person objects.
        var persons = LoadCsv(csvPath);

        // Ensure output directory exists.
        string outputDir = "output";
        Directory.CreateDirectory(outputDir);

        // Generate a separate document for each CSV record.
        foreach (var person in persons)
        {
            // Load the template fresh for each document.
            var doc = new Document(templatePath);
            var engine = new ReportingEngine();

            // Build the report using the current person as the root object.
            engine.BuildReport(doc, person, "record");

            // Save the generated document.
            string outputPath = Path.Combine(outputDir, $"Report_{person.Name}.docx");
            doc.Save(outputPath);
        }
    }

    // Simple data model matching the CSV columns.
    public class Person
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }
        public string City { get; set; } = "";
    }

    // Reads the CSV file and returns a list of Person objects.
    private static List<Person> LoadCsv(string path)
    {
        var list = new List<Person>();
        using var reader = new StreamReader(path);
        // Read header line.
        string? headerLine = reader.ReadLine();
        if (headerLine == null)
            return list;

        while (!reader.EndOfStream)
        {
            string? line = reader.ReadLine();
            if (string.IsNullOrWhiteSpace(line))
                continue;

            string[] parts = line.Split(',');
            if (parts.Length < 3)
                continue;

            var person = new Person
            {
                Name = parts[0].Trim(),
                Age = int.TryParse(parts[1].Trim(), out int age) ? age : 0,
                City = parts[2].Trim()
            };
            list.Add(person);
        }

        return list;
    }
}
