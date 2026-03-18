using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables; // Added for Table type

namespace AsposeWordsWhereExample
{
    // External type whose static property will be used inside the LINQ Where clause.
    public static class ExternalHelper
    {
        // This property can be changed at runtime to affect the filtering logic.
        public static int MinAge { get; set; } = 30;
    }

    // Simple data class that will be filtered.
    public class Person
    {
        // Made nullable to satisfy the non‑nullable warning (or could use required).
        public string Name { get; set; } = null!;
        public int Age { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Prepare a sample collection of Person objects.
            List<Person> people = new List<Person>
            {
                new Person { Name = "Alice", Age = 25 },
                new Person { Name = "Bob",   Age = 35 },
                new Person { Name = "Carol", Age = 28 },
                new Person { Name = "David", Age = 42 }
            };

            // Create a ReportingEngine instance.
            ReportingEngine engine = new ReportingEngine();

            // Register the external type so that its members can be used in templates or expressions.
            // In this example we register it to demonstrate the concept; the LINQ query uses it directly.
            engine.KnownTypes.Add(typeof(ExternalHelper));

            // Use the LINQ Where extension method with a lambda that references the external static property.
            // This filters the collection to only include persons whose Age is greater than ExternalHelper.MinAge.
            IEnumerable<Person> filtered = people.Where(p => p.Age > ExternalHelper.MinAge);

            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a heading.
            builder.Writeln("People older than " + ExternalHelper.MinAge + ":");
            builder.Writeln();

            // Insert a simple table to display the filtered results.
            Table table = builder.StartTable(); // Now Table type is recognized.
            // Header row.
            builder.InsertCell();
            builder.Write("Name");
            builder.InsertCell();
            builder.Write("Age");
            builder.EndRow();

            // Populate rows with filtered data.
            foreach (Person person in filtered)
            {
                builder.InsertCell();
                builder.Write(person.Name);
                builder.InsertCell();
                builder.Write(person.Age.ToString());
                builder.EndRow();
            }

            builder.EndTable();

            // Save the document to disk.
            doc.Save("FilteredPeople.docx");
        }
    }
}
