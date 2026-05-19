using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Data model with first and last name properties.
    public class Person
    {
        public string FirstName { get; set; } = "";
        public string LastName { get; set; } = "";
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a LINQ Reporting tag that concatenates first and last name with a space.
            builder.Writeln("Full Name: <<[person.FirstName + \" \" + person.LastName]>>");

            // Prepare sample data.
            Person person = new Person
            {
                FirstName = "John",
                LastName = "Doe"
            };

            // Build the report using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, person, "person");

            // Save the generated report.
            doc.Save("Report.docx");
        }
    }
}
