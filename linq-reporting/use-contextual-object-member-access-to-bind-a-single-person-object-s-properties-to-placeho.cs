using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model with public properties.
    public class Person
    {
        public string FirstName { get; set; } = string.Empty;
        public string LastName { get; set; } = string.Empty;
        public int Age { get; set; }

        public Person(string firstName, string lastName, int age)
        {
            FirstName = firstName;
            LastName = lastName;
            Age = age;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a sample Person instance.
            Person person = new Person("John", "Doe", 30);

            // Create a blank document and insert LINQ Reporting tags.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Use contextual object member access: the root name is "person".
            builder.Writeln("Name: <<[person.FirstName]>> <<[person.LastName]>>");
            builder.Writeln("Age: <<[person.Age]>>");

            // Build the report by binding the Person object to the template.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, person, "person");

            // Save the generated report.
            doc.Save("Report.docx");
        }
    }
}
