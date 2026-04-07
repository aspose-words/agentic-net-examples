using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using System.Text;

Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

public class Person
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
    public string Status { get; set; } = string.Empty;
}

public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
}

public partial class Program
{
    public static void Main()
    {
        // 1. Prepare sample CSV data.
        string csvPath = "people.csv";
        File.WriteAllLines(csvPath, new[]
        {
            "Name,Age,Status",
            "Alice,30,Active",
            "Bob,45,Inactive",
            "Charlie,25,Active",
            "Diana,40,Inactive"
        });

        // 2. Load CSV into a list of Person objects.
        List<Person> allPersons = LoadCsv(csvPath);

        // 3. Filter rows where Status == "Active" using LINQ Where.
        List<Person> activePersons = allPersons
            .Where(p => string.Equals(p.Status, "Active", StringComparison.OrdinalIgnoreCase))
            .ToList();

        // 4. Prepare the model for the reporting engine.
        ReportModel model = new()
        {
            Persons = activePersons
        };

        // 5. Create the template document programmatically.
        string templatePath = "template.docx";
        CreateTemplate(templatePath);

        // 6. Load the template (ensuring it is fully loaded before building the report).
        Document doc = new Document(templatePath);

        // 7. Build the report using the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // 8. Save the final report.
        string outputPath = "Report.docx";
        doc.Save(outputPath);
    }

    private static List<Person> LoadCsv(string path)
    {
        var persons = new List<Person>();
        string[] lines = File.ReadAllLines(path);
        if (lines.Length < 2)
            return persons; // No data.

        // Assume first line contains headers.
        for (int i = 1; i < lines.Length; i++)
        {
            string line = lines[i];
            if (string.IsNullOrWhiteSpace(line))
                continue;

            string[] parts = line.Split(',');
            if (parts.Length != 3)
                continue; // Skip malformed lines.

            if (!int.TryParse(parts[1], out int age))
                age = 0;

            persons.Add(new Person
            {
                Name = parts[0].Trim(),
                Age = age,
                Status = parts[2].Trim()
            });
        }

        return persons;
    }

    private static void CreateTemplate(string path)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Active Persons Report");
        builder.Writeln("----------------------");

        // LINQ Reporting foreach tag iterating over the Persons collection.
        builder.Writeln("<<foreach [person in Persons]>>");
        builder.Writeln("Name: <<[person.Name]>>, Age: <<[person.Age]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(path);
    }
}
