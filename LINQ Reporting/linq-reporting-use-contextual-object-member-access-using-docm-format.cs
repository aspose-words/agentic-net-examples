using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsReportingDemo
{
    // Sample data class whose members will be accessed from the template.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Load a DOCM template that contains LINQ Reporting tags.
            // Example tag in the template: <<[person].Name>> and <<[person].Age>>
            Document doc = new Document("Template.docm");

            // Create the data source object.
            var person = new Person
            {
                Name = "John Doe",
                Age = 30
            };

            // Initialize the reporting engine.
            ReportingEngine engine = new ReportingEngine();

            // Build the report using the overload that allows referencing the data source object itself.
            // The third argument ("person") is the name used inside the template to refer to the object.
            engine.BuildReport(doc, person, "person");

            // Save the populated document as a DOCM file.
            doc.Save("Report.docm");
        }
    }
}
