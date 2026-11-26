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
        // Register code page provider for possible legacy encodings.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample CSV data with European date format (dd/MM/yyyy).
        string csvPath = "data.csv";
        File.WriteAllText(csvPath, "Name,Date\r\nAlice,31/12/2023\r\nBob,15/01/2024");

        // Load CSV data using a custom culture (en-GB) to parse dates correctly.
        var persons = new List<Person>();
        var culture = new CultureInfo("en-GB");
        using (var reader = new StreamReader(csvPath))
        {
            // Read header.
            string? headerLine = reader.ReadLine();
            if (headerLine == null) throw new InvalidOperationException("CSV file is empty.");

            // Read data lines.
            while (!reader.EndOfStream)
            {
                string? line = reader.ReadLine();
                if (string.IsNullOrWhiteSpace(line)) continue;

                string[] parts = line.Split(',');
                if (parts.Length != 2) continue;

                string name = parts[0].Trim();
                string dateString = parts[1].Trim();

                // Parse date using the custom culture.
                DateTime date = DateTime.Parse(dateString, culture);
                persons.Add(new Person { Name = name, Date = date });
            }
        }

        // Create a simple Word template with LINQ Reporting tags.
        string templatePath = "template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        builder.Writeln("Report of Persons:");
        builder.Writeln("<<foreach [person in Persons]>>");
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Date: <<[person.Date]>>");
        builder.Writeln("<</foreach>>");

        doc.Save(templatePath);

        // Load the template for reporting.
        var template = new Document(templatePath);

        // Prepare the root model.
        var model = new ReportModel { Persons = persons };

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(template, model, "model");

        // Save the generated report.
        string outputPath = "report.docx";
        template.Save(outputPath);
    }
}

public class Person
{
    public string Name { get; set; } = string.Empty;
    public DateTime Date { get; set; }
}

public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
}
