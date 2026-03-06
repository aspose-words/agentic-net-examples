using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsReportingDemo
{
    // Simple data entity that matches the fields used in the template.
    public class Person
    {
        public string FullName { get; set; }
        public string Address { get; set; }

        public Person(string fullName, string address)
        {
            FullName = fullName;
            Address = address;
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the plain‑text source file.
            // Expected format per line: FullName|Address
            const string txtPath = @"Data\People.txt";

            // Load the TXT file, split each line and convert to a collection of Person objects.
            List<Person> people = LoadPeopleFromTxt(txtPath);

            // Load the template document that contains reporting tags, e.g. <<foreach [persons]>><<FullName>> - <<Address>><</foreach>>
            Document template = new Document(@"Templates\PeopleReportTemplate.docx");

            // Build the report using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            // The data source name "persons" must match the name used in the template.
            engine.BuildReport(template, people, "persons");

            // Save the generated report.
            template.Save(@"Output\PeopleReport.docx");
        }

        // Reads a TXT file and converts each line into a Person object.
        private static List<Person> LoadPeopleFromTxt(string filePath)
        {
            var result = new List<Person>();

            // Ensure the file exists before attempting to read.
            if (!File.Exists(filePath))
                throw new FileNotFoundException($"The data file '{filePath}' was not found.");

            // Read all lines, ignoring empty ones.
            foreach (string line in File.ReadAllLines(filePath))
            {
                if (string.IsNullOrWhiteSpace(line))
                    continue;

                // Split the line by the pipe character.
                string[] parts = line.Split('|');
                if (parts.Length != 2)
                    throw new FormatException($"Invalid line format: '{line}'. Expected 'FullName|Address'.");

                string fullName = parts[0].Trim();
                string address = parts[1].Trim();

                result.Add(new Person(fullName, address));
            }

            return result;
        }
    }
}
