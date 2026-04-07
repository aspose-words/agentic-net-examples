using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;
using System.Text;

namespace AsposeWordsLinqReportingExample
{
    // Data model representing a person record.
    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
        public string City { get; set; } = string.Empty;
        public decimal Salary { get; set; }
    }

    // Wrapper class used as the root data source for the report.
    public class ReportModel
    {
        public List<Person> Persons { get; set; } = new();
    }

    class Program
    {
        static void Main()
        {
            // Register code page provider for any encoding needs (e.g., JSON files).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Step 1: Create sample JSON data file.
            string jsonPath = "people.json";
            var sampleData = new List<Person>
            {
                new Person { Name = "Alice", Age = 28, City = "New York", Salary = 72000m },
                new Person { Name = "Bob",   Age = 35, City = "Chicago",   Salary = 54000m },
                new Person { Name = "Carol", Age = 42, City = "New York", Salary = 88000m },
                new Person { Name = "Dave",  Age = 31, City = "Los Angeles", Salary = 61000m },
                new Person { Name = "Eve",   Age = 45, City = "New York", Salary = 45000m }
            };
            File.WriteAllText(jsonPath, JsonConvert.SerializeObject(sampleData, Formatting.Indented));

            // Step 2: Load JSON data into objects.
            var allPersons = JsonConvert.DeserializeObject<List<Person>>(File.ReadAllText(jsonPath)) ?? new List<Person>();

            // Step 3: Apply a complex predicate using LINQ Where.
            // Keep persons older than 30, living in New York, with a salary of at least 60,000.
            var filteredPersons = allPersons
                .Where(p => p.Age > 30 && p.City == "New York" && p.Salary >= 60000m)
                .ToList();

            // Step 4: Create a Word template programmatically.
            var templatePath = "template.docx";
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Insert a heading.
            builder.Writeln("Filtered Persons Report");
            builder.Writeln();

            // Insert a foreach block that iterates over the filtered collection.
            builder.Writeln("<<foreach [person in Persons]>>");
            builder.Writeln("Name: <<[person.Name]>>");
            builder.Writeln("Age: <<[person.Age]>>");
            builder.Writeln("City: <<[person.City]>>");
            builder.Writeln("Salary: $<<[person.Salary]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            doc.Save(templatePath);

            // Step 5: Load the template for reporting.
            var reportDoc = new Document(templatePath);

            // Prepare the root data source with the filtered list.
            var model = new ReportModel { Persons = filteredPersons };

            // Step 6: Build the report using Aspose.Words ReportingEngine.
            var engine = new ReportingEngine();
            engine.BuildReport(reportDoc, model, "model");

            // Step 7: Save the generated report.
            string outputPath = "report.docx";
            reportDoc.Save(outputPath);

            // Inform that the process completed (no interactive input required).
            Console.WriteLine($"Report generated successfully: {Path.GetFullPath(outputPath)}");
        }
    }
}
