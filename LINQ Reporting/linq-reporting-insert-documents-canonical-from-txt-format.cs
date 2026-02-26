using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsReportingExample
{
    // Simple data class that matches the fields used in the template.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Paths to the input files and the output document.
            string txtPath = @"C:\Data\people.txt";          // Example: "John Doe,30"
            string templatePath = @"C:\Templates\ReportTemplate.docx";
            string outputPath = @"C:\Output\ReportResult.docx";

            // Load the plain‑text file, split each line by a comma and project it into a list of Person objects.
            List<Person> people = File.ReadAllLines(txtPath)
                .Where(line => !string.IsNullOrWhiteSpace(line))
                .Select(line =>
                {
                    string[] parts = line.Split(',');
                    return new Person
                    {
                        Name = parts[0].Trim(),
                        Age = int.Parse(parts[1].Trim())
                    };
                })
                .ToList();

            // Load the Word template document.
            Document doc = new Document(templatePath);

            // Create the reporting engine and populate the template with the data source.
            ReportingEngine engine = new ReportingEngine();
            // The data source name ("data") must match the name used in the template tags, e.g. <<foreach [in data]>>
            engine.BuildReport(doc, people, "data");

            // Save the populated document.
            doc.Save(outputPath);
        }
    }
}
