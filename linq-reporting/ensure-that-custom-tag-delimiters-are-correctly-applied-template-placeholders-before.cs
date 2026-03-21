using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Replacing;

namespace CustomTagDelimiterExample
{
    public class Person
    {
        public string FirstName { get; set; }
        public string LastName { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Create an in‑memory template containing custom delimiters {{ }}.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);
            builder.Writeln("{{FirstName}} {{LastName}}");

            // Convert custom delimiters to the format expected by ReportingEngine (<<[ ]>>).
            template.Range.Replace("{{", "<<[", new FindReplaceOptions());
            template.Range.Replace("}}", "]>>", new FindReplaceOptions());

            // Prepare the data source.
            Person person = new Person { FirstName = "John", LastName = "Doe" };

            // Configure and execute the reporting engine.
            ReportingEngine engine = new ReportingEngine
            {
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };
            engine.BuildReport(template, person, "person");

            // Save the generated report.
            template.Save("GeneratedReport.docx");
        }
    }
}
