using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExample
{
    // Simple data model with first and last name.
    public class Person
    {
        public string FirstName { get; set; } = string.Empty;
        public string LastName { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a LINQ Reporting expression that concatenates first and last name with a space.
            // The expression is evaluated during BuildReport.
            builder.Writeln("<<[person.FirstName + \" \" + person.LastName]>>");

            // Prepare sample data.
            Person person = new Person
            {
                FirstName = "John",
                LastName = "Doe"
            };

            // Build the report using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            // The root object name in the template is "person", so we pass the name explicitly.
            engine.BuildReport(doc, person, "person");

            // Save the generated document.
            doc.Save("Report.docx");
        }
    }
}
