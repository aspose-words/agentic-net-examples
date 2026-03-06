using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReport
{
    // Simple POCO class that will be used as a LINQ data source.
    public class Person
    {
        public string FirstName { get; set; }
        public string LastName  { get; set; }
        public int Age          { get; set; }

        public Person(string firstName, string lastName, int age)
        {
            FirstName = firstName;
            LastName  = lastName;
            Age       = age;
        }
    }

    class Program
    {
        static void Main()
        {
            // Load the DOCX template that contains Reporting Engine tags, e.g. <<[persons.FirstName]>>.
            Document template = new Document("Template.docx");

            // Prepare a LINQ-friendly data source – a list of Person objects.
            List<Person> people = new List<Person>
            {
                new Person("John",  "Doe",   30),
                new Person("Jane",  "Smith", 25),
                new Person("Alice", "Brown", 28)
            };

            // Create the ReportingEngine instance.
            ReportingEngine engine = new ReportingEngine();

            // Build the report. The third argument is the name used in the template to reference the data source.
            // In the template you would use tags like <<[persons.FirstName]>> or <<foreach [persons]>><<[FirstName]>>...
            engine.BuildReport(template, people, "persons");

            // Save the generated report to a new DOCX file.
            template.Save("ReportResult.docx");
        }
    }
}
