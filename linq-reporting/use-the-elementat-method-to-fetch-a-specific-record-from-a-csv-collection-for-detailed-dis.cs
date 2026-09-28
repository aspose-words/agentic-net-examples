using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    public int Id { get; set; }
    public string Name { get; set; } = "";
    public int Age { get; set; }
}

public class Program
{
    public static void Main()
    {
        // Ensure code page provider for possible CSV encoding.
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // 1. Create sample CSV file.
        string csvPath = "people.csv";
        File.WriteAllText(csvPath,
@"Id,Name,Age
1,John Doe,30
2,Jane Smith,25
3,Bob Johnson,40
4,Alice Brown,22");

        // 2. Load CSV data into a list of Person objects.
        List<Person> people = new List<Person>();
        foreach (var line in File.ReadAllLines(csvPath).Skip(1)) // Skip header
        {
            var parts = line.Split(',');
            if (parts.Length != 3) continue;
            if (!int.TryParse(parts[0], out int id)) continue;
            var name = parts[1];
            if (!int.TryParse(parts[2], out int age)) continue;
            people.Add(new Person { Id = id, Name = name, Age = age });
        }

        // 3. Use ElementAt to fetch the third record (index 2).
        Person selectedPerson = people.ElementAt(2); // Bob Johnson

        // 4. Create a Word template with LINQ Reporting tags.
        string templatePath = "template.docx";
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Detailed Person Report");
        builder.Writeln("<<[person.Id]>>");
        builder.Writeln("<<[person.Name]>>");
        builder.Writeln("<<[person.Age]>>");
        templateDoc.Save(templatePath);

        // 5. Load the template and build the report.
        Document reportDoc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, selectedPerson, "person");

        // 6. Save the generated report.
        string outputPath = "PersonReport.docx";
        reportDoc.Save(outputPath);

        // Optional: indicate completion (no interactive wait).
        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}
