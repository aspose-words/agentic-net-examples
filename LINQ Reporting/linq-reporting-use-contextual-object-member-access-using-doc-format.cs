using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data class that will be used as the data source.
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

        // Example of a contextual member that can be accessed directly from the template.
        public string FullName => $"{FirstName} {LastName}";
    }

    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Write a template that accesses the data source object itself (contextual access).
            // The data source will be referenced by the name "person".
            // <<[person.FullName]>> accesses a member of the data source object.
            // <<[person.Age]>> accesses another member.
            builder.Writeln("Name: <<[person.FullName]>>");
            builder.Writeln("Age : <<[person.Age]>>");

            // Prepare the data source.
            Person person = new Person("John", "Doe", 42);

            // Create the reporting engine.
            ReportingEngine engine = new ReportingEngine();

            // Build the report using the overload that allows referencing the data source object itself.
            // The third argument is the name that will be used inside the template.
            engine.BuildReport(doc, person, "person");

            // Save the populated document.
            doc.Save("LinqReporting_ContextualAccess.docx");
        }
    }
}
