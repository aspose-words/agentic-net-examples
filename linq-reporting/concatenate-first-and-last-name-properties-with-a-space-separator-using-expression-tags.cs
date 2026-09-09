using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExample
{
    // Data model with first and last name.
    public class Person
    {
        public string FirstName { get; set; } = string.Empty;
        public string LastName { get; set; } = string.Empty;
    }

    public class Program
    {
        public static void Main()
        {
            // Create a blank document and a builder to insert the template.
            var doc = new Document();
            var builder = new DocumentBuilder(doc);

            // Insert a LINQ Reporting expression that concatenates first and last name with a space.
            builder.Writeln("<<[person.FirstName + \" \" + person.LastName]>>");

            // Prepare the data source.
            var person = new Person
            {
                FirstName = "John",
                LastName = "Doe"
            };

            // Build the report using the ReportingEngine.
            var engine = new ReportingEngine();
            engine.BuildReport(doc, person, "person");

            // Save the generated report.
            doc.Save("Report.docx");
        }
    }
}
