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

            // ---------------------------------------------------------------------------
            // Configure the ReportingEngine.
            //
            // 1. Restrict the original mutable type (Person) so that its members cannot be
            //    accessed from templates. This prevents any accidental use of setters.
            // 2. Build the report using the read‑only wrapper type (PersonReadOnly) which
            //    only exposes getters.
            // ---------------------------------------------------------------------------
            ReportingEngine.SetRestrictedTypes(typeof(Person));

            // Create the engine instance. No special options are required for this scenario.
            var engine = new ReportingEngine();

            // Load a template document that contains reporting tags, e.g.:
            //   <<foreach [in readOnlyPeople]>>
            //       <<[Name]>> is <<[Age]>>
            //   <</foreach>>
            var doc = new Document("Template.docx");

            // Build the report using the read‑only collection as the data source.
            engine.BuildReport(doc, new { readOnlyPeople });

            // Save the resulting document.
            doc.Save("Report_Output.docx");
        }
    }
}
