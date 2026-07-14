using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model with a collection of persons.
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
            var model = new ReportModel();
            model.Persons.Add(new Person { Name = "Alice", Age = 30 });
            model.Persons.Add(new Person { Name = "Bob", Age = 45 });

            // Build a template document programmatically.
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Insert a foreach loop that iterates over the Persons collection.
            builder.Writeln("<<foreach [p in Persons]>>");
            builder.Writeln("Name: <<[p.Name]>>, Age: <<[p.Age]>>");
            builder.Writeln("<</foreach>>");

            // Configure the reporting engine to inline error messages.
            var engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.InlineErrorMessages;

            // Build the report and capture the success flag.
            bool success = engine.BuildReport(doc, model, "model");

            // Output the success flag to the console.
            Console.WriteLine($"Report build success: {success}");

            // Save the generated report.
            doc.Save("ReportOutput.docx");
        }
    }
}
