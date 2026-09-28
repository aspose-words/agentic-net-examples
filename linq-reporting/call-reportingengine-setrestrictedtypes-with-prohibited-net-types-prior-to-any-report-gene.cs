using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingRestrictedTypesExample
{
    // Simple data model used as the root object for the report.
    public class ReportModel
    {
        public List<Person> Persons { get; set; } = new();
    }

    public class Person
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare sample data.
            var model = new ReportModel
            {
                Persons = new List<Person>
                {
                    new Person { Name = "Alice", Age = 30 },
                    new Person { Name = "Bob", Age = 25 },
                    new Person { Name = "Charlie", Age = 35 }
                }
            };

            // Create a template document programmatically.
            var templatePath = "Template.docx";
            var builder = new DocumentBuilder();
            builder.Writeln("<<foreach [p in Persons]>>");
            builder.Writeln("Name: <<[p.Name]>>, Age: <<[p.Age]>>");
            builder.Writeln("<</foreach>>");
            builder.Document.Save(templatePath);

            // Load the template document.
            var doc = new Document(templatePath);

            // Restrict prohibited .NET types before building the report.
            ReportingEngine.SetRestrictedTypes(new[]
            {
                typeof(System.IO.FileInfo),
                typeof(System.Diagnostics.Process)
            });

            // Build the report.
            var engine = new ReportingEngine();
            engine.BuildReport(doc, model, "model");

            // Save the generated report.
            var outputPath = "Report.docx";
            doc.Save(outputPath);
        }
    }
}
