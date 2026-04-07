using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    // Simple data model used by the LINQ Reporting template.
    public class Person
    {
        // Name will be displayed in the report.
        public string Name { get; set; } = string.Empty;

        // Optional is intentionally left null to produce an empty paragraph after processing.
        public string? Optional { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Create a sample data object.
            Person person = new Person
            {
                Name = "John Doe",
                Optional = null // This will result in an empty paragraph in the template.
            };

            // -----------------------------------------------------------------
            // 1. Build the template document programmatically.
            // -----------------------------------------------------------------
            Document doc = new Document();                     // Create a blank document.
            DocumentBuilder builder = new DocumentBuilder(doc); // Attach a builder.

            // Paragraph that will become empty after the tag is evaluated (Optional is null).
            builder.Writeln("<<[person.Optional]>>");

            // Paragraph that contains actual data.
            builder.Writeln("Hello <<[person.Name]>>!");

            // -----------------------------------------------------------------
            // 2. Configure the ReportingEngine to remove empty paragraphs.
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

            // -----------------------------------------------------------------
            // 3. Build the report using the data source.
            // -----------------------------------------------------------------
            // The root object name used in the template is "person".
            engine.BuildReport(doc, person, "person");

            // -----------------------------------------------------------------
            // 4. Save the generated report.
            // -----------------------------------------------------------------
            doc.Save("Report.docx");
        }
    }
}
