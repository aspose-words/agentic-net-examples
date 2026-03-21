// Program.cs
using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace ReportingExample
{
    // Define a data class that contains both getters and setters.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }

    // Define a read‑only wrapper that exposes only getter properties.
    // Templates will be able to read these members but cannot invoke any setters.
    public class PersonReadOnly
    {
        private readonly Person _inner;
        public PersonReadOnly(Person inner) => _inner = inner;
        public string Name => _inner.Name;
        public int Age => _inner.Age;
    }

    class Program
    {
        static void Main()
        {
            // Prepare sample data.
            var people = new List<Person>
            {
                new Person { Name = "Alice", Age = 30 },
                new Person { Name = "Bob",   Age = 25 }
            };

            // Convert the list to a list of read‑only wrappers.
            var readOnlyPeople = people.Select(p => new PersonReadOnly(p)).ToList();

            // Restrict the mutable type so its members cannot be accessed from templates.
            ReportingEngine.SetRestrictedTypes(typeof(Person));

            // Create the engine instance.
            var engine = new ReportingEngine();

            // Create a simple template document in‑memory.
            var doc = new Document();
            var builder = new DocumentBuilder(doc);
            builder.Writeln("<<foreach [in readOnlyPeople]>>");
            builder.Writeln("Name: <<[Name]>>, Age: <<[Age]>>");
            builder.Writeln("<</foreach>>");

            // Build the report using the read‑only collection as the data source.
            engine.BuildReport(doc, readOnlyPeople, "readOnlyPeople");

            // Save the resulting document.
            doc.Save("Report_Output.docx");
        }
    }
}
