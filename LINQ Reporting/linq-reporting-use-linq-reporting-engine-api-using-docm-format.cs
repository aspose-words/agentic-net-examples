using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Simple data class to be used as a data source.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Path to the DOCM template that contains LINQ Reporting tags, e.g. <<[persons.Name]>>.
            string templatePath = @"C:\Templates\ReportTemplate.docm";

            // Load the DOCM template.
            Document doc = new Document(templatePath);

            // Prepare sample data.
            List<Person> persons = new List<Person>
            {
                new Person { Name = "Alice", Age = 30 },
                new Person { Name = "Bob", Age = 45 },
                new Person { Name = "Charlie", Age = 28 }
            };

            // Create the reporting engine.
            ReportingEngine engine = new ReportingEngine();

            // Build the report. The data source name "persons" must match the name used in the template.
            engine.BuildReport(doc, persons, "persons");

            // Save the populated document as DOCM.
            string outputPath = @"C:\Output\GeneratedReport.docm";
            doc.Save(outputPath);
        }
    }
}
