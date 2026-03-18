using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsFilteringExample
{
    // Simple POCO representing a person with an Age property.
    public class Person
    {
        public string First { get; set; }
        public string Last  { get; set; }
        public int    Age   { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Load the Word template that contains the reporting tags, e.g. <<foreach [persons]>><<[First]>><</foreach>>.
            Document template = new Document("Template.docx");

            // Create a source collection of persons.
            List<Person> persons = new List<Person>
            {
                new Person { First = "John",  Last = "Doe",   Age = 25 },
                new Person { First = "Jane",  Last = "Smith", Age = 16 },
                new Person { First = "Bob",   Last = "Brown", Age = 30 }
            };

            // Filter the collection to include only adults (Age > 18).
            List<Person> adults = persons.Where(p => p.Age > 18).ToList();

            // Populate the template using the ReportingEngine.
            // The data source name "persons" matches the tag used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(template, adults, "persons");

            // Save the populated document.
            template.Save("Report.docx");
        }
    }
}
