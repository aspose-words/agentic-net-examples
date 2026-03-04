using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsReportingEngineDemo
{
    // Simple POCO class that will be used as a data source for the report.
    public class Person
    {
        public string FirstName { get; set; }
        public string LastName  { get; set; }
        public int    Age       { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Path to the macro‑enabled template (DOTM) that contains LINQ Reporting Engine tags.
            const string templatePath = @"C:\Templates\ReportTemplate.dotm";

            // Load the DOTM template document.
            Document template = new Document(templatePath);

            // Prepare sample data.
            var people = new List<Person>
            {
                new Person { FirstName = "John",  LastName = "Doe",   Age = 30 },
                new Person { FirstName = "Jane",  LastName = "Smith", Age = 25 },
                new Person { FirstName = "Alice", LastName = "Brown", Age = 28 }
            };

            // Create the reporting engine.
            ReportingEngine engine = new ReportingEngine
            {
                // Example option: remove paragraphs that become empty after processing.
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };

            // Build the report. The data source name ("people") can be referenced in the template.
            engine.BuildReport(template, people, "people");

            // Save the populated document as a macro‑enabled template (DOTM).
            const string outputPath = @"C:\Output\GeneratedReport.dotm";
            template.Save(outputPath, SaveFormat.Dotm);

            Console.WriteLine("Report generated successfully at: " + outputPath);
        }
    }
}
