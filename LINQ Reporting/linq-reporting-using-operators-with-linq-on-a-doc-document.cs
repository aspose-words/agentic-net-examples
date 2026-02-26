using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Simple POCO class representing a person.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a blank Word document that will serve as the template.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Insert a heading.
            builder.Writeln("People older than 30:");

            // Insert a LINQ Reporting Engine foreach tag.
            // The tag syntax <<foreach [person]>>...<</foreach>> will be processed by ReportingEngine.
            builder.Writeln("<<foreach [person]>>");
            builder.Writeln("Name: <<[person.Name]>>");
            builder.Writeln("Age: <<[person.Age]>>");
            builder.Writeln("<</foreach>>");

            // 2. Prepare a data source – a list of Person objects.
            List<Person> allPeople = new List<Person>
            {
                new Person { Name = "Alice",   Age = 28 },
                new Person { Name = "Bob",     Age = 35 },
                new Person { Name = "Charlie", Age = 42 },
                new Person { Name = "Diana",   Age = 22 }
            };

            // 3. Apply LINQ operators to the data source.
            // For example, filter out people whose age is 30 or less.
            IEnumerable<Person> filteredPeople = allPeople
                .Where(p => p.Age > 30)          // Only keep people older than 30.
                .OrderBy(p => p.Name);           // Sort alphabetically by name.

            // 4. Build the report using ReportingEngine.
            ReportingEngine engine = new ReportingEngine();

            // The data source name ("person") must match the name used in the template tags.
            // BuildReport will replace the foreach block with rows for each Person in filteredPeople.
            engine.BuildReport(template, filteredPeople, "person");

            // 5. Save the resulting document.
            template.Save("LinqReportingResult.docx");
        }
    }
}
