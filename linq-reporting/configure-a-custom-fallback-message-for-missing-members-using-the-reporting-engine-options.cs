using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data model with only Name property.
    public class Person
    {
        public string Name { get; set; } = "John Doe";
        // Note: Age property is intentionally missing to demonstrate fallback.
    }

    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert LINQ Reporting tags. The Age tag refers to a missing member.
            builder.Writeln("Name: <<[person.Name]>>");
            builder.Writeln("Age: <<[person.Age]>>"); // Age does not exist in Person.

            // Prepare the reporting engine.
            ReportingEngine engine = new ReportingEngine();
            // Enable handling of missing members.
            engine.Options = ReportBuildOptions.AllowMissingMembers;
            // Set custom fallback message for missing members.
            engine.MissingMemberMessage = "N/A";

            // Create the data source.
            Person person = new Person();

            // Build the report. The root object name must match the tag prefix.
            engine.BuildReport(doc, person, "person");

            // Save the generated report.
            doc.Save("Report.docx");
        }
    }
}
