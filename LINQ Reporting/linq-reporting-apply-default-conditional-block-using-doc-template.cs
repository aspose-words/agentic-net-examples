using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data source class used in the template.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }

        // Property that may be missing in some records to demonstrate default handling.
        public string Occupation { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Load the DOC template that contains a conditional block.
            // Example template syntax:
            // <<if [ds.Age >= 18]>><<[ds.Name]>> is an adult.<</if>>
            // <<else>><<[ds.Name]>> is a minor.<</else>>
            Document template = new Document("Template.docx");

            // Prepare a data source with a missing member (Occupation) to trigger the default block.
            Person[] people = new[]
            {
                new Person { Name = "Alice", Age = 25, Occupation = "Engineer" },
                new Person { Name = "Bob",   Age = 16 } // Occupation is missing.
            };

            // Configure the ReportingEngine.
            ReportingEngine engine = new ReportingEngine
            {
                // Allow missing members so the engine does not throw an exception.
                Options = ReportBuildOptions.AllowMissingMembers | ReportBuildOptions.RemoveEmptyParagraphs,
                // Text to display when a missing member is encountered.
                MissingMemberMessage = "N/A"
            };

            // Build the report using the data source.
            // The data source name "ds" is referenced in the template.
            engine.BuildReport(template, people, "ds");

            // Save the generated report.
            template.Save("ReportOutput.docx");
        }
    }
}
