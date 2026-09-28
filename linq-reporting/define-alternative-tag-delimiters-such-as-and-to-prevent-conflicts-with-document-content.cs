using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingCustomDelimiters
{
    // Sample data model.
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

            // Create a template document programmatically.
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Use alternative delimiters [[ and ]] for LINQ Reporting tags.
            builder.Writeln("[[foreach [p in Persons]]]");
            builder.Writeln("Name: [[[p.Name]]]  Age: [[[p.Age]]]");
            builder.Writeln("[[</foreach]]]");

            // Save the template (optional, for inspection).
            doc.Save("Template.docx");

            // Configure the reporting engine.
            var engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None; // No special options needed.

            // Build the report using the custom delimiters.
            engine.BuildReport(doc, model, "model");

            // Save the generated report.
            doc.Save("Report_Output.docx");
        }
    }
}
