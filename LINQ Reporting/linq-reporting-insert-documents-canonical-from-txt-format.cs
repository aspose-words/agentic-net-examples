using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsReportingExample
{
    // Simple POCO that matches the fields used in the Word template.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public string City { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Path to the plain‑text data file (each line: Name|Age|City).
            string txtFilePath = @"C:\Data\people.txt";

            // Path to the Word template that contains tags like <<[person.Name]>>, <<[person.Age]>>, etc.
            string templatePath = @"C:\Templates\PeopleReport.docx";

            // Path where the generated report will be saved.
            string outputPath = @"C:\Reports\PeopleReportResult.docx";

            // Load the TXT file, parse each line with LINQ and create a list of Person objects.
            List<Person> people = File.ReadAllLines(txtFilePath)
                .Where(line => !string.IsNullOrWhiteSpace(line))
                .Select(line =>
                {
                    // Expected format: Name|Age|City
                    string[] parts = line.Split('|');
                    return new Person
                    {
                        Name = parts.Length > 0 ? parts[0].Trim() : string.Empty,
                        Age = parts.Length > 1 && int.TryParse(parts[1].Trim(), out int age) ? age : 0,
                        City = parts.Length > 2 ? parts[2].Trim() : string.Empty
                    };
                })
                .ToList();

            // Load the template document (create rule is used internally by the Document constructor).
            Document templateDoc = new Document(templatePath);

            // Build the report using the ReportingEngine.
            // The data source name "person" must match the tags used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(templateDoc, people, "person");

            // Save the populated document (save rule).
            templateDoc.Save(outputPath);
        }
    }
}
