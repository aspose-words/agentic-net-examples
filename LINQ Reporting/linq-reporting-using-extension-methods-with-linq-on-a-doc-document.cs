using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Data model used in the template.
    public class Person
    {
        public string Name { get; set; }
        public DateTime BirthDate { get; set; }
    }

    // Extension methods that will be called from the template via LINQ Reporting.
    public static class PersonExtensions
    {
        // Returns the age of a person based on the current date.
        public static int Age(this Person person)
        {
            var today = DateTime.Today;
            var age = today.Year - person.BirthDate.Year;
            if (person.BirthDate > today.AddYears(-age)) age--;
            return age;
        }

        // Returns true if the person is an adult (18+).
        public static bool IsAdult(this Person person) => person.Age() >= 18;
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a blank document that will serve as the template.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Insert a simple LINQ Reporting template that iterates over a collection
            // and uses the extension methods defined above.
            builder.Writeln("People Report");
            builder.Writeln("==============");
            // foreach over the collection named 'people'
            builder.Writeln("<<foreach [people]>>");
            // Output each person's name, age (extension method) and adult flag (extension method)
            builder.Writeln("Name: <<[Name]>>");
            builder.Writeln("Age: <<[Age()]>>");
            builder.Writeln("Adult: <<[IsAdult()]>>");
            builder.Writeln("<</foreach>>");

            // 2. Prepare the data source.
            List<Person> people = new List<Person>
            {
                new Person { Name = "Alice",   BirthDate = new DateTime(1990, 5, 12) },
                new Person { Name = "Bob",     BirthDate = new DateTime(2005, 3, 30) },
                new Person { Name = "Charlie", BirthDate = new DateTime(1982, 11, 5) }
            };

            // 3. Configure the ReportingEngine.
            ReportingEngine engine = new ReportingEngine
            {
                // Allow missing members so that the engine does not throw if a member is not found.
                Options = ReportBuildOptions.AllowMissingMembers
            };
            // Register the static class that contains the extension methods.
            engine.KnownTypes.Add(typeof(PersonExtensions));

            // 4. Build the report. The data source name 'people' matches the name used in the template.
            engine.BuildReport(template, people, "people");

            // 5. Save the generated document.
            template.Save("PeopleReport.docx");
        }
    }
}
