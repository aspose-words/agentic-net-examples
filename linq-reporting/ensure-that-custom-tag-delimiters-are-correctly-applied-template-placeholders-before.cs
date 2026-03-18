using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Replacing;

namespace CustomTagDelimiterExample
{
    // Sample data source class
    public class Person
    {
        public string FirstName { get; set; }
        public string LastName { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Load the template document that contains custom delimiters (e.g., {{FirstName}})
            Document template = new Document("TemplateWithCustomDelimiters.docx");

            // Replace custom delimiters with the delimiters expected by ReportingEngine (<<[ ... ]>>)
            // Example: {{FirstName}} -> <<[person.FirstName]>>
            // This simple replacement assumes that the placeholder name does not contain the delimiters themselves.
            // Adjust the patterns as needed for more complex scenarios.

            // Replace opening delimiter "{{" with "<<["
            template.Range.Replace("{{", "<<[", new FindReplaceOptions());

            // Replace closing delimiter "}}" with "]>>"
            template.Range.Replace("}}", "]>>", new FindReplaceOptions());

            // Prepare the data source
            Person person = new Person { FirstName = "John", LastName = "Doe" };

            // Configure the reporting engine
            ReportingEngine engine = new ReportingEngine
            {
                // Remove empty paragraphs that may appear after replacement
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };

            // Build the report using the modified template and the data source.
            // The data source name ("person") must match the name used in the template placeholders.
            engine.BuildReport(template, person, "person");

            // Save the generated report
            template.Save("GeneratedReport.docx");
        }
    }
}
