using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Simple POCO class that will be used as a data source for the report.
    public class Person
    {
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public int Age { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Path to the template document that contains LINQ Reporting Engine tags.
            // Example tag in the template: <<[person.FirstName]>> <<[person.LastName]>> (Age: <<[person.Age]>>)
            string templatePath = @"C:\Templates\PersonReportTemplate.docx";

            // Path where the generated report will be saved.
            string outputPath = @"C:\Reports\PersonReport.docx";

            // Load the template document (lifecycle rule: load).
            Document doc = new Document(templatePath);

            // Create a data source instance.
            Person person = new Person
            {
                FirstName = "John",
                LastName = "Doe",
                Age = 30
            };

            // Initialize the ReportingEngine (no special options required here).
            ReportingEngine engine = new ReportingEngine();

            // Build the report by populating the template with the data source.
            // The third parameter is the name used to reference the data source inside the template.
            engine.BuildReport(doc, person, "person");

            // Save the populated document (lifecycle rule: save).
            doc.Save(outputPath);
        }
    }
}
