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
        // Register code pages for CSV encoding.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Create a sample CSV file.
        string csvPath = Path.Combine(outputDir, "data.csv");
        File.WriteAllText(csvPath, "Id,Name,Age\r\n1,John Doe,30\r\n2,Jane Smith,25\r\n");

        // Load CSV data into a strongly‑typed list.
        List<Person> persons = LoadCsv(csvPath);

        // Create a Word template with LINQ Reporting tags.
        string templatePath = Path.Combine(outputDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Customer List:");
        builder.Writeln("<<foreach [row in Data]>>");
        builder.Writeln("Id: <<[row.Id]>>, Name: <<[row.Name]>>, Age: <<[row.Age]>>");
        builder.Writeln("<</foreach>>");
        templateDoc.Save(templatePath);

        // Load the template.
        Document doc = new Document(templatePath);

        // Disable reflection optimization for this import.
        ReportingEngine.UseReflectionOptimization = false;

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        bool success = engine.BuildReport(doc, persons, "Data");

        // Save the generated report.
        string reportPath = Path.Combine(outputDir, "report.docx");
        doc.Save(reportPath);

        // Indicate completion.
        Console.WriteLine($"Report generated at: {reportPath} (Success: {success})");
    }

    private static List<Person> LoadCsv(string path)
    {
        var list = new List<Person>();
        string[] lines = File.ReadAllLines(path);
        // Skip header line.
        for (int i = 1; i < lines.Length; i++)
        {
            string line = lines[i];
            if (string.IsNullOrWhiteSpace(line))
                continue;

            string[] parts = line.Split(',');
            if (parts.Length >= 3 &&
                int.TryParse(parts[0], out int id) &&
                int.TryParse(parts[2], out int age))
            {
                list.Add(new Person
                {
                    Id = id,
                    Name = parts[1],
                    Age = age
                });
            }
        }
        return list;
    }

    public class Person
    {
        public int Id { get; set; } = 0;
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; } = 0;
    }
}
