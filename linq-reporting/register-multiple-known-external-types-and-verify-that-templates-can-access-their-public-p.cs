using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Sample data model classes.
    public class Person
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }
        public Address Address { get; set; } = new();
    }

    public class Address
    {
        public string City { get; set; } = "";
        public string Country { get; set; } = "";
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
            // Ensure the output folder exists.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);

            // 1. Create a template document with LINQ Reporting tags.
            string templatePath = Path.Combine(outputDir, "Template.docx");
            CreateTemplate(templatePath);

            // 2. Prepare sample data.
            ReportModel model = new ReportModel
            {
                Persons = new List<Person>
                {
                    new Person
                    {
                        Name = "Alice",
                        Age = 30,
                        Address = new Address { City = "New York", Country = "USA" }
                    },
                    new Person
                    {
                        Name = "Bob",
                        Age = 45,
                        Address = new Address { City = "London", Country = "UK" }
                    }
                }
            };

            // 3. Load the template.
            Document templateDoc = new Document(templatePath);

            // 4. Register external types so the template can access their public members without reflection.
            ReportingEngine engine = new ReportingEngine();
            engine.KnownTypes.Add(typeof(Person));
            engine.KnownTypes.Add(typeof(Address));

            // 5. Build the report.
            engine.BuildReport(templateDoc, model, "model");

            // 6. Save the generated report.
            string reportPath = Path.Combine(outputDir, "Report.docx");
            templateDoc.Save(reportPath);

            // Indicate completion.
            Console.WriteLine($"Report generated at: {reportPath}");
        }

        // Helper method to create the template document.
        private static void CreateTemplate(string filePath)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a heading.
            builder.Writeln("Person Report");
            builder.Writeln();

            // Begin a foreach loop over the Persons collection.
            builder.Writeln("<<foreach [p in Persons]>>");
            builder.Writeln("Name: <<[p.Name]>>");
            builder.Writeln("Age: <<[p.Age]>>");
            builder.Writeln("City: <<[p.Address.City]>>");
            builder.Writeln("Country: <<[p.Address.Country]>>");
            builder.Writeln("<</foreach>>");

            // Save the template.
            doc.Save(filePath);
        }
    }
}
