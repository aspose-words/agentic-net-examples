using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for CSV encoding support.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample CSV data.
        string csvPath = "data.csv";
        File.WriteAllText(csvPath, "Name,Age\r\nAlice,30\r\nBob,25\r\nCharlie,35");

        // Load CSV data into a strongly‑typed collection.
        var persons = new List<Person>();
        foreach (var line in File.ReadAllLines(csvPath))
        {
            // Skip header line.
            if (line.StartsWith("Name", StringComparison.OrdinalIgnoreCase))
                continue;

            var parts = line.Split(',');
            if (parts.Length >= 2 &&
                int.TryParse(parts[1], out int age))
            {
                persons.Add(new Person
                {
                    Name = parts[0],
                    Age = age
                });
            }
        }

        // Create a Word template with LINQ Reporting tags.
        string templatePath = "template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("People Report");
        builder.Writeln("<<foreach [person in csvData]>>");
        builder.Writeln("Name: <<[person.Name]>>, Age: <<[person.Age]>>");
        builder.Writeln("<</foreach>>");
        doc.Save(templatePath);

        // Load the template.
        var template = new Document(templatePath);

        // Build the report using the strongly‑typed collection as the data source.
        var engine = new ReportingEngine();
        engine.BuildReport(template, persons, "csvData");

        // Save the generated report.
        string outputPath = "output.docx";
        template.Save(outputPath);
    }
}
