using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExample
{
    // Business object representing a person.
    public class Person
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }
    }

    // Wrapper model that will be passed as the root data source.
    public class ReportModel
    {
        public List<Person> Persons { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Create sample data.
            var model = new ReportModel
            {
                Persons = new List<Person>
                {
                    new Person { Name = "Alice", Age = 30 },
                    new Person { Name = "Bob", Age = 25 },
                    new Person { Name = "Charlie", Age = 35 }
                }
            };

            // Create a template document with LINQ Reporting tags.
            var templatePath = "Template.docx";
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            builder.Writeln("People Report");
            builder.Writeln("<<foreach [p in Persons]>>");
            builder.Writeln("Name: <<[p.Name]>>, Age: <<[p.Age]>>");
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
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
