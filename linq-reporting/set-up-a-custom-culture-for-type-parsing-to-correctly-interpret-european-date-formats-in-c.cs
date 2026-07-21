using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Ensure the working directory exists.
        Directory.CreateDirectory("Output");

        // 1. Create sample CSV data with European date format (dd/MM/yyyy).
        string csvPath = Path.Combine("Output", "data.csv");
        File.WriteAllText(csvPath,
            "Name,BirthDate\r\n" +
            "Alice,31/12/1990\r\n" +
            "Bob,15/07/1985\r\n" +
            "Charlie,01/01/2000\r\n");

        // 2. Set a custom culture that interprets dates in European format.
        CultureInfo europeanCulture = new CultureInfo("fr-FR"); // dd/MM/yyyy pattern
        CultureInfo.DefaultThreadCurrentCulture = europeanCulture;
        CultureInfo.DefaultThreadCurrentUICulture = europeanCulture;

        // 3. Load CSV and parse into model objects using the custom culture.
        List<Person> persons = new();
        using (var reader = new StreamReader(csvPath))
        {
            // Skip header.
            if (!reader.EndOfStream) reader.ReadLine();

            while (!reader.EndOfStream)
            {
                string line = reader.ReadLine();
                if (string.IsNullOrWhiteSpace(line)) continue;

                string[] parts = line.Split(',');
                if (parts.Length != 2) continue;

                string name = parts[0].Trim();
                string dateString = parts[1].Trim();

                // Parse date using the custom culture.
                if (!DateTime.TryParse(dateString, europeanCulture, DateTimeStyles.None, out DateTime birthDate))
                {
                    // Fallback to exact format if needed.
                    birthDate = DateTime.ParseExact(dateString, "dd/MM/yyyy", europeanCulture);
                }

                persons.Add(new Person { Name = name, BirthDate = birthDate });
            }
        }

        // 4. Prepare the data model for the report.
        ReportModel model = new() { Persons = persons };

        // 5. Create a Word template programmatically with LINQ Reporting tags.
        string templatePath = Path.Combine("Output", "template.docx");
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);

        builder.Writeln("People Report");
        builder.Writeln("");
        builder.Writeln("<<foreach [person in Persons]>>");
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Birth Date: <<[person.BirthDate]>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // 6. Load the template and generate the report.
        Document reportDoc = new(templatePath);
        ReportingEngine engine = new();
        engine.BuildReport(reportDoc, model, "model");

        // 7. Save the generated report.
        string reportPath = Path.Combine("Output", "report.docx");
        reportDoc.Save(reportPath);
    }
}

public class Person
{
    public string Name { get; set; } = string.Empty;
    public DateTime BirthDate { get; set; }
}

public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
}
