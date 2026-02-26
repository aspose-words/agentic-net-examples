using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data entity.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Prepare a list of persons.
            List<Person> people = new List<Person>
            {
                new Person { Name = "Alice", Age = 30 },
                new Person { Name = "Bob", Age = 17 },
                new Person { Name = "Charlie", Age = 22 }
            };

            // Use LINQ to select only adults (age >= 18).
            List<Person> adults = people.Where(p => p.Age >= 18).ToList();

            // Create a template document with Reporting Engine tags.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // The tags <<[person.Name]>> and <<[person.Age]>> will be replaced by the engine.
            builder.Writeln("Name: <<[person.Name]>>");
            builder.Writeln("Age: <<[person.Age]>>");
            builder.Writeln(); // Add an empty line between records.

            // Build the report using the filtered data source.
            ReportingEngine engine = new ReportingEngine();
            // The data source name "person" matches the tags used in the template.
            engine.BuildReport(doc, adults, "person");

            // Save the resulting document.
            doc.Save("LinqReportingResult.docx");
        }
    }
}
