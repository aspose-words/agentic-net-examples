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
        // Register code page provider (required for some encodings)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample CSV data with European date format (dd/MM/yyyy)
        string csvPath = "people.csv";
        File.WriteAllText(csvPath, "Name;BirthDate\nJohn Doe;15/04/1990\nJane Smith;23/12/1985", Encoding.UTF8);

        // Create a custom culture that parses dates in the European format
        CultureInfo europeanCulture = (CultureInfo)CultureInfo.InvariantCulture.Clone();
        europeanCulture.DateTimeFormat.ShortDatePattern = "dd/MM/yyyy";
        europeanCulture.DateTimeFormat.DateSeparator = "/";

        // Load CSV data using the custom culture
        List<Person> persons = new();
        using (var reader = new StreamReader(csvPath))
        {
            // Skip header
            string? header = reader.ReadLine();
            while (!reader.EndOfStream)
            {
                string? line = reader.ReadLine();
                if (string.IsNullOrWhiteSpace(line))
                    continue;

                string[] parts = line.Split(';');
                if (parts.Length != 2)
                    continue;

                string name = parts[0];
                string dateText = parts[1];

                DateTime birthDate = DateTime.Parse(dateText, europeanCulture);
                persons.Add(new Person { Name = name, BirthDate = birthDate });
            }
        }

        // Create the report model
        ReportModel model = new() { Persons = persons };

        // Build the template document programmatically
        string templatePath = "Template.docx";
        Document templateDoc = new();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("People Report");
        builder.Writeln("<<foreach [person in Persons]>>");
        builder.Writeln("Name: <<[person.Name]>>, Birth Date: <<[person.BirthDate]>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // Load the template for reporting
        Document reportDoc = new(templatePath);

        // Build the report
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report
        string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
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
