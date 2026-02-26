using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExample
{
    // Simple data entity used in the report.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }

    // Wrapper class that holds the collection referenced from the template.
    public class ReportData
    {
        public List<Person> Persons { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Path to the DOCX template that contains LINQ Reporting tags,
            // e.g. <<foreach [Persons]>><<[Name]>> <<[Age]>><</foreach>>.
            string templatePath = "Template.docx";

            // Path where the generated report will be saved.
            string outputPath = "Result.docx";

            // Load the template document (lifecycle rule: load).
            Document template = new Document(templatePath);

            // Prepare the data source.
            ReportData data = new ReportData
            {
                Persons = GetSamplePersons()
            };

            // Create the reporting engine and populate the template (feature rule: BuildReport).
            ReportingEngine engine = new ReportingEngine();
            // The third parameter (data source name) can be null or empty when the object itself
            // does not need to be referenced directly in the template.
            engine.BuildReport(template, data, null);

            // Save the populated document (lifecycle rule: save).
            template.Save(outputPath);
        }

        // Generates a list of sample persons to demonstrate sequential output.
        private static List<Person> GetSamplePersons()
        {
            return new List<Person>
            {
                new Person { Name = "Alice Johnson", Age = 28 },
                new Person { Name = "Bob Smith", Age = 35 },
                new Person { Name = "Carol Davis", Age = 42 }
            };
        }
    }
}
