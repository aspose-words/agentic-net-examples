using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model that matches the CSV columns.
    public class Person
    {
        public string FirstName { get; set; } = string.Empty;
        public string LastName  { get; set; } = string.Empty;
        public string Email     { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Ensure the output folder exists.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
            Directory.CreateDirectory(outputDir);

            // 1. Create a sample CSV file.
            string csvPath = Path.Combine(Directory.GetCurrentDirectory(), "people.csv");
            CreateSampleCsv(csvPath);

            // 2. Read the CSV file into a list of Person objects.
            List<Person> persons = LoadPersonsFromCsv(csvPath);

            // 3. Create a single‑row template programmatically and save it.
            string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "template.docx");
            CreateTemplateDocument(templatePath);

            // 4. For each person generate an individual document.
            int index = 1;
            foreach (Person person in persons)
            {
                // Load the template.
                Document doc = new Document(templatePath);

                // Build the report using the person as the root object.
                ReportingEngine engine = new ReportingEngine
                {
                    Options = ReportBuildOptions.None
                };
                // The template tags reference the root name "record".
                engine.BuildReport(doc, person, "record");

                // Save the generated document.
                string outputPath = Path.Combine(outputDir, $"Person_{index}.docx");
                doc.Save(outputPath);
                index++;
            }
        }

        // Creates a simple CSV file with a header row and a few sample records.
        private static void CreateSampleCsv(string path)
        {
            string[] lines =
            {
                "FirstName,LastName,Email",
                "John,Doe,john.doe@example.com",
                "Jane,Smith,jane.smith@example.com",
                "Bob,Johnson,bob.johnson@example.com"
            };
            File.WriteAllLines(path, lines);
        }

        // Parses the CSV file into a list of Person objects.
        private static List<Person> LoadPersonsFromCsv(string path)
        {
            var persons = new List<Person>();
            string[] allLines = File.ReadAllLines(path);
            if (allLines.Length < 2)
                return persons; // No data.

            // Assume first line contains headers.
            for (int i = 1; i < allLines.Length; i++)
            {
                string line = allLines[i];
                if (string.IsNullOrWhiteSpace(line))
                    continue;

                string[] parts = line.Split(',');
                if (parts.Length != 3)
                    continue; // Skip malformed lines.

                persons.Add(new Person
                {
                    FirstName = parts[0].Trim(),
                    LastName  = parts[1].Trim(),
                    Email     = parts[2].Trim()
                });
            }
            return persons;
        }

        // Builds a template document that contains LINQ Reporting tags.
        private static void CreateTemplateDocument(string path)
        {
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // The root object name used in BuildReport will be "record".
            builder.Writeln("<<[record.FirstName]>> <<[record.LastName]>>");
            builder.Writeln("Email: <<[record.Email]>>");

            // Save the template.
            template.Save(path);
        }
    }
}
