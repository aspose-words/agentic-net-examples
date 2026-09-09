using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingCsvFilter
{
    // Simple data model representing a CSV row.
    public class Person
    {
        public string Name { get; set; } = "";
        public string Status { get; set; } = "";
    }

    // Wrapper class used as the root object for the LINQ Reporting engine.
    public class ReportModel
    {
        public List<Person> persons { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some encodings).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare sample CSV data.
            string csvPath = "people.csv";
            File.WriteAllLines(csvPath, new[]
            {
                "Name,Status",
                "Alice,Active",
                "Bob,Inactive",
                "Charlie,Active",
                "Diana,Inactive"
            });

            // Load CSV rows into a list of Person objects.
            List<Person> allPersons = File.ReadAllLines(csvPath)
                .Skip(1) // Skip header.
                .Select(line => line.Split(','))
                .Where(parts => parts.Length == 2)
                .Select(parts => new Person
                {
                    Name = parts[0].Trim(),
                    Status = parts[1].Trim()
                })
                .ToList();

            // Filter rows where Status equals "Active" using LINQ Where.
            List<Person> activePersons = allPersons
                .Where(p => string.Equals(p.Status, "Active", StringComparison.OrdinalIgnoreCase))
                .ToList();

            // Create the report template programmatically.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Insert a foreach tag that iterates over the filtered collection.
            builder.Writeln("<<foreach [p in persons]>>");
            builder.Writeln("Name: <<[p.Name]>> | Status: <<[p.Status]>>");
            builder.Writeln("<</foreach>>");

            // Save the template (optional, demonstrates lifecycle rule).
            string templatePath = "template.docx";
            template.Save(templatePath);

            // Load the template back (simulating a separate load step).
            Document doc = new Document(templatePath);

            // Prepare the root model with the filtered data.
            ReportModel model = new ReportModel
            {
                persons = activePersons
            };

            // Build the report using the LINQ Reporting engine.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the final report.
            string outputPath = "Report_ActivePersons.docx";
            doc.Save(outputPath);
        }
    }
}
