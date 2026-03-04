using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsReportingEngineDemo
{
    // Simple data class that will be used as a data source for the report.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Load the DOCM template that contains LINQ Reporting Engine tags.
            // The template file should exist at the specified path.
            Document template = new Document("Template.docm");

            // Prepare the data source – a list of Person objects.
            List<Person> people = new List<Person>
            {
                new Person { Name = "Alice", Age = 30 },
                new Person { Name = "Bob",   Age = 45 },
                new Person { Name = "Carol", Age = 27 }
            };

            // Initialize the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();

            // Build the report by populating the template with the data source.
            // The third argument ("people") is the name used to reference the data source in the template.
            engine.BuildReport(template, people, "people");

            // Save the populated document as a DOCM file.
            template.Save("ReportOutput.docm");
        }
    }
}
