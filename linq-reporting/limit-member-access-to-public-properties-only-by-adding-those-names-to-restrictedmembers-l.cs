using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model with public properties.
    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }

        // Private member that should not be accessible from the template.
        private string Secret { get; set; } = "TopSecret";

        // Public method to retrieve the secret (not used in the template).
        public string GetSecret() => Secret;
    }

    // Wrapper class required because ReportingEngine does not accept anonymous types.
    public class ReportModel
    {
        public List<Person> Persons { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare sample data.
            var persons = new List<Person>
            {
                new Person { Name = "Alice", Age = 30 },
                new Person { Name = "Bob", Age = 45 },
                new Person { Name = "Charlie", Age = 28 }
            };

            // Create the template document programmatically.
            var templatePath = "template.docx";
            CreateTemplate(templatePath);

            // Load the template.
            var doc = new Document(templatePath);

            // Build the report using the LINQ Reporting engine.
            var engine = new ReportingEngine();

            // Wrap the data source in a public class so that BuildReport accepts it.
            var model = new ReportModel { Persons = persons };
            engine.BuildReport(doc, model, "model");

            // Save the generated report.
            var outputPath = "report.docx";
            doc.Save(outputPath);

            Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
        }

        // Creates a simple Word template with LINQ Reporting tags.
        private static void CreateTemplate(string filePath)
        {
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Begin a foreach loop over the collection "Persons".
            builder.Writeln("<<foreach [p in Persons]>>");
            // Insert public property values.
            builder.Writeln("Name: <<[p.Name]>>");
            builder.Writeln("Age: <<[p.Age]>>");
            // End the foreach loop.
            builder.Writeln("<</foreach>>");

            // Save the template.
            doc.Save(filePath);
        }
    }
}
