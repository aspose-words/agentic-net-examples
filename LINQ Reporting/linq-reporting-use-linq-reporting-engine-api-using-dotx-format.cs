using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingDemo
{
    // Simple POCO class that will be used as the data source for the report.
    public class Person
    {
        public string FirstName { get; set; }
        public string LastName  { get; set; }
        public int    Age       { get; set; }

        public Person(string firstName, string lastName, int age)
        {
            FirstName = firstName;
            LastName  = lastName;
            Age       = age;
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the DOTX template that contains LINQ Reporting Engine tags.
            // Example tag inside the template: <<[person.FirstName]>> <<[person.LastName]>>
            string templatePath = @"C:\Templates\PersonReport.dotx";

            // Load the DOTX template into a Document object.
            Document doc = new Document(templatePath);

            // Create a data source instance.
            Person person = new Person("John", "Doe", 42);

            // Initialize the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();

            // Build the report by populating the template with the data source.
            // The third argument is the name used to reference the data source inside the template.
            engine.BuildReport(doc, person, "person");

            // Save the generated report as a DOCX file.
            string outputPath = @"C:\Reports\PersonReport.docx";
            doc.Save(outputPath);
        }
    }
}
