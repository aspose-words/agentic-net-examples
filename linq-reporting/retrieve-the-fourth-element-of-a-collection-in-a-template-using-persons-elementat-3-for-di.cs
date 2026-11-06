using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using System.Text;

namespace LinqReportingExample
{
    // Simple data model for a person.
    public class Person
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }
    }

    // Wrapper model that holds a collection of persons.
    public class ReportModel
    {
        public List<Person> Persons { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Register code page provider (required for some Aspose.Words features).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Prepare sample data with at least four persons.
            var model = new ReportModel
            {
                Persons = new List<Person>
                {
                    new() { Name = "Alice", Age = 30 },
                    new() { Name = "Bob", Age = 25 },
                    new() { Name = "Charlie", Age = 28 },
                    new() { Name = "Diana", Age = 32 },   // Fourth element (index 3)
                    new() { Name = "Ethan", Age = 22 }
                }
            };

            // Create a template document programmatically.
            var templatePath = "Template.docx";
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            builder.Writeln("LINQ Reporting Example - Fourth Person");
            // Use LINQ Reporting tag to retrieve the fourth element via ElementAt(3).
            builder.Writeln("Name: <<[model.Persons.ElementAt(3).Name]>>");
            builder.Writeln("Age: <<[model.Persons.ElementAt(3).Age]>>");

            // Save the template.
            doc.Save(templatePath);

            // Load the template for report generation.
            var reportDoc = new Document(templatePath);

            // Build the report using the LINQ Reporting engine.
            var engine = new ReportingEngine();
            engine.BuildReport(reportDoc, model, "model");

            // Save the generated report.
            var outputPath = "Report.docx";
            reportDoc.Save(outputPath);
        }
    }
}
