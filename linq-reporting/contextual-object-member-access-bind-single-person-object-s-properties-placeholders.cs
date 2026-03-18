using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Bibliography;

namespace AsposeWordsExample
{
    class Program
    {
        static void Main()
        {
            // Load the template document that contains placeholders such as <<[person.First]>>, <<[person.Middle]>>, <<[person.Last]>>.
            Document template = new Document("Template.docx");

            // Create a single Person instance.
            // Constructor parameters: last name, first name, middle name.
            Person person = new Person("Doe", "John", "M.");

            // Initialize the reporting engine.
            ReportingEngine engine = new ReportingEngine();

            // Build the report using the overload that allows referencing the data source object itself.
            // The third argument ("person") is the name used in the template to access the object's members.
            engine.BuildReport(template, person, "person");

            // Save the populated document.
            template.Save("Result.docx");
        }
    }
}
