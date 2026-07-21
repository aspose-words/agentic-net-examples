using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV parsing (if needed).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample CSV data.
        string csvPath = "people.csv";
        File.WriteAllLines(csvPath, new[]
        {
            "Id,Name,Status",
            "1,John Doe,Active",
            "2,Jane Smith,Inactive",
            "3,Bob Johnson,Active"
        });

        // Load CSV rows into a list of Person objects.
        List<Person> allPeople = LoadPeopleFromCsv(csvPath);

        // Filter rows where Status == "Active".
        List<Person> activePeople = allPeople.Where(p => p.Status == "Active").ToList();

        // Create a LINQ Reporting template programmatically.
        string templatePath = "template.docx";
        CreateTemplate(templatePath);

        // Load the template document.
        Document doc = new Document(templatePath);

        // Prepare the model with filtered data.
        ReportModel model = new ReportModel { People = activePeople };

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save("report.docx");
    }

    // Parses the CSV file into a list of Person objects.
    private static List<Person> LoadPeopleFromCsv(string path)
    {
        var lines = File.ReadAllLines(path);
        var people = new List<Person>();

        // Assume first line contains headers.
        for (int i = 1; i < lines.Length; i++)
        {
            var parts = lines[i].Split(',');
            if (parts.Length >= 3 &&
                int.TryParse(parts[0], out int id))
            {
                people.Add(new Person
                {
                    Id = id,
                    Name = parts[1],
                    Status = parts[2]
                });
            }
        }

        return people;
    }

    // Creates a simple template that iterates over People collection.
    private static void CreateTemplate(string path)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("<<foreach [person in People]>>");
        builder.Writeln("Id: <<[person.Id]>>, Name: <<[person.Name]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(path);
    }
}

// Data model exposed to the template.
public class ReportModel
{
    public List<Person> People { get; set; } = new();
}

// Represents a row from the CSV file.
public class Person
{
    public int Id { get; set; }
    public string Name { get; set; } = string.Empty;
    public string Status { get; set; } = string.Empty;
}
