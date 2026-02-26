using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model containing an array.
    public class PersonData
    {
        // Array of anonymous objects (could be any type).
        public Person[] Persons { get; set; }

        public PersonData()
        {
            // Initialize the array with sample data.
            Persons = new[]
            {
                new Person { Name = "Alice", Age = 30 },
                new Person { Name = "Bob", Age = 25 },
                new Person { Name = "Charlie", Age = 35 }
            };
        }
    }

    // POCO representing a single record.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // 1. Load the WORDML (DOCX) template that contains LINQ Reporting tags.
            //    Example tag in the template: <<foreach [ds.Persons]>><<[Name]>> (Age: <<[Age]>>)<</foreach>>
            Document template = new Document("Template.docx");

            // 2. Create the data source. The property Persons is an array.
            //    ReportingEngine can work directly with arrays because they implement IEnumerable,
            //    but to demonstrate conversion to a canonical collection we explicitly cast to List<T>.
            PersonData dataSource = new PersonData();

            // Convert the array to a List<Person> – this is the canonical collection type
            // that the reporting engine will iterate over.
            List<Person> personsList = new List<Person>(dataSource.Persons);

            // 3. Build the report. We pass the list as a named data source ("ds").
            ReportingEngine engine = new ReportingEngine();
            // The anonymous object wrapper allows us to expose the list under a property name.
            var wrapper = new { Persons = personsList };
            engine.BuildReport(template, wrapper, "ds");

            // 4. Save the populated document.
            template.Save("ReportResult.docx");
        }
    }
}
