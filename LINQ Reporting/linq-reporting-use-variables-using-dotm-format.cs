using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data class used as a data source for the LINQ reporting engine.
    public class Person
    {
        public string FirstName { get; set; }
        public string LastName  { get; set; }
        public int Age          { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Load a DOTM template that contains LINQ Reporting tags.
            // Example template content (saved as Template.dotm):
            //   <<[DocVar]>>                     // displays a document variable
            //   <<[person.FirstName]>> <<[person.LastName]>> (Age: <<[person.Age]>>)
            Document doc = new Document("Template.dotm");

            // Add a document variable that can be referenced from the template.
            // The variable can be displayed using a DOCVARIABLE field or directly via <<[DocVar]>> syntax.
            doc.Variables.Add("DocVar", "Report generated on " + DateTime.Now.ToString("yyyy-MM-dd"));

            // Prepare the data source for the report.
            var people = new List<Person>
            {
                new Person { FirstName = "John",  LastName = "Doe",   Age = 30 },
                new Person { FirstName = "Jane",  LastName = "Smith", Age = 25 },
                new Person { FirstName = "Alice", LastName = "Brown", Age = 28 }
            };

            // Create the reporting engine.
            ReportingEngine engine = new ReportingEngine();

            // Build the report.
            // The third argument ("person") is the name used in the template to reference the data source object.
            // This allows the template to access both the object's members (e.g., person.FirstName)
            // and the object itself (e.g., <<[person]>>).
            engine.BuildReport(doc, people, "person");

            // Save the populated document.
            doc.Save("Report.docx");
        }
    }
}
