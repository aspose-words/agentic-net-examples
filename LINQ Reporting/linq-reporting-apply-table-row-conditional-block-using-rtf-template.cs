using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingConditionalRow
{
    // Simple data model used as the data source for the report.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }

        // When true the corresponding table row will be shown, otherwise it will be removed.
        public bool ShowRow { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Load the RTF template that contains a table with a conditional block.
            // The template should have a row like:
            // <<if [ds.ShowRow]>>
            //   <<[ds.Name]>>
            //   <<[ds.Age]>>
            // <<endif>>
            Document doc = new Document(@"C:\Templates\ConditionalRowTemplate.rtf");

            // Prepare a list of persons – this will be the data source.
            List<Person> people = new List<Person>
            {
                new Person { Name = "Alice", Age = 30, ShowRow = true },
                new Person { Name = "Bob",   Age = 45, ShowRow = false }, // This row will be omitted.
                new Person { Name = "Carol", Age = 27, ShowRow = true }
            };

            // Create the reporting engine.
            ReportingEngine engine = new ReportingEngine
            {
                // Remove empty paragraphs that may appear after conditional rows are removed.
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };

            // Build the report. The data source name "ds" is used inside the template.
            engine.BuildReport(doc, people, "ds");

            // Save the populated document.
            doc.Save(@"C:\Output\ConditionalRowResult.docx");
        }
    }
}
