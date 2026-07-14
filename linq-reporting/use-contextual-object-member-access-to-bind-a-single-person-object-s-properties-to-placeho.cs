using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model representing a person.
    public class Person
    {
        public string FirstName { get; set; } = "";
        public string LastName { get; set; } = "";
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
            // Create a blank document that will serve as the template.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Insert LINQ Reporting tags that reference the Person object's members.
            builder.Writeln("Name: <<[person.FirstName]>> <<[person.LastName]>>");
            builder.Writeln("Age: <<[person.Age]>>");

            // Prepare the data source – a single Person instance.
            Person person = new Person("John", "Doe", 30);

            // Build the report using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            // The root object name used in the template is "person".
            engine.BuildReport(template, person, "person");

            // Save the generated report.
            template.Save("Report.docx");
        }
    }
}
