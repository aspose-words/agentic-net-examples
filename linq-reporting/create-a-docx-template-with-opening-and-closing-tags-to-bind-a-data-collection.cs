using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExample
{
    // Simple data model with a collection to bind.
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
            // Register code page provider (required for some data sources).
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            // Prepare sample data.
            var model = new ReportModel
            {
                Persons = new List<Person>
                {
                    new() { Name = "Alice", Age = 30 },
                    new() { Name = "Bob", Age = 45 },
                    new() { Name = "Charlie", Age = 28 }
                }
            };

            // Create a DOCX template with LINQ Reporting tags.
            const string templatePath = "Template.docx";
            var templateDoc = new Document();
            var builder = new DocumentBuilder(templateDoc);

            builder.Writeln("People List:");
            builder.Writeln("<<foreach [person in Persons]>>");
            builder.Writeln("Name: <<[person.Name]>>, Age: <<[person.Age]>>");
            builder.Writeln("<</foreach>>");

            templateDoc.Save(templatePath);

            // Load the template for report generation.
            var reportDoc = new Document(templatePath);

            // Build the report using the LINQ Reporting engine.
            var engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.None; // default options
            engine.BuildReport(reportDoc, model); // root object is model; tags reference its members directly

            // Save the generated report.
            const string reportPath = "Report.docx";
            reportDoc.Save(reportPath);
        }
    }
}
