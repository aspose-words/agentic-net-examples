using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Data model classes
    public class Person
    {
        public string Name { get; set; } = "";
        public int Age { get; set; }
    }

    public class ReportModel
    {
        public List<Person> Persons { get; set; } = new();
    }

    public class Program
    {
        public static void Main()
        {
            // Create a template document with LINQ Reporting tags.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Begin a foreach loop over the Persons collection.
            builder.Writeln("<<foreach [p in Persons]>>");
            // Correct tag – will be replaced with the person's name.
            builder.Writeln("Name: <<[p.Name]>>");
            // Incorrect tag – property 'Agee' does not exist, will trigger an inline error message.
            builder.Writeln("Age: <<[p.Agee]>>");
            // End the foreach loop.
            builder.Writeln("<</foreach>>");

            // Save the template to disk.
            const string templatePath = "template.docx";
            template.Save(templatePath);

            // Load the template for report generation.
            Document report = new Document(templatePath);

            // Prepare sample data.
            var model = new ReportModel
            {
                Persons = new List<Person>
                {
                    new Person { Name = "Alice", Age = 30 },
                    new Person { Name = "Bob", Age = 25 }
                }
            };

            // Configure the reporting engine to inline error messages.
            ReportingEngine engine = new ReportingEngine
            {
                Options = ReportBuildOptions.InlineErrorMessages
            };

            // Build the report. The returned flag indicates whether parsing succeeded.
            bool success = engine.BuildReport(report, model, "model");

            // Save the generated report.
            const string outputPath = "output.docx";
            report.Save(outputPath);

            // Output the success flag.
            Console.WriteLine($"Report build success: {success}");
            Console.WriteLine($"Template saved to: {templatePath}");
            Console.WriteLine($"Report saved to: {outputPath}");
        }
    }
}
