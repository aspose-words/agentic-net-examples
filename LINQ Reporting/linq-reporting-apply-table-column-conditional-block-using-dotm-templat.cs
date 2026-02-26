using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReporting
{
    class Program
    {
        static void Main()
        {
            // Load the DOTM template that contains a table with a conditional column.
            Document template = new Document("Template.dotm");

            // Prepare a data source – a list of objects.
            // The property ShowAge will be used in the template to conditionally display the Age column.
            List<Person> data = new List<Person>
            {
                new Person { Name = "John Doe", Age = 30, ShowAge = true },
                new Person { Name = "Jane Smith", Age = 25, ShowAge = false },
                new Person { Name = "Bob Johnson", Age = 40, ShowAge = true }
            };

            // Create the reporting engine.
            ReportingEngine engine = new ReportingEngine();

            // Optional: remove empty paragraphs that may appear after conditional blocks are removed.
            engine.Options = ReportBuildOptions.RemoveEmptyParagraphs;

            // Build the report. The data source name "people" must match the name used in the template
            // (e.g., <<foreach [people]>><<[Name]>> ... <<if [people.ShowAge]>> <<[Age]>> <<endif>> <<endforeach>>).
            engine.BuildReport(template, data, "people");

            // Save the populated document.
            template.Save("Report.docx");
        }
    }

    // Simple POCO class used as the data source.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
        // Determines whether the Age column should be shown for this record.
        public bool ShowAge { get; set; }
    }
}
