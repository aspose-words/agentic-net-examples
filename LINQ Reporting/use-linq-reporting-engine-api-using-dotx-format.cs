using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsReportingDemo
{
    // Simple data class used as the data source for the report.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Load the DOTX template document.
            Document template = new Document("Template.dotx");

            // Create an instance of the data source.
            Person data = new Person
            {
                Name = "John Doe",
                Age = 30
            };

            // Initialize the LINQ Reporting Engine.
            ReportingEngine engine = new ReportingEngine();

            // Populate the template with data. The data source name ("person") can be used in the template.
            engine.BuildReport(template, data, "person");

            // Save the populated document as a DOTX file.
            template.Save("Report.dotx", SaveFormat.Dotx);
        }
    }
}
