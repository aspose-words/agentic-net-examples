using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingDemo
{
    // Simple data model that will be used as the data source for the LINQ Reporting Engine.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Load an HTML file that contains the template markup.
            //    The Document constructor automatically detects the format (HTML).
            // -----------------------------------------------------------------
            string htmlTemplatePath = @"C:\Templates\PersonTemplate.html";
            Document templateDoc = new Document(htmlTemplatePath);

            // -----------------------------------------------------------------
            // 2. Prepare the data source – a list of Person objects.
            //    The ReportingEngine can work with any non‑dynamic .NET type.
            // -----------------------------------------------------------------
            List<Person> people = new List<Person>
            {
                new Person { Name = "Alice Johnson", Age = 29 },
                new Person { Name = "Bob Smith",     Age = 42 },
                new Person { Name = "Carol Lee",     Age = 35 }
            };

            // -----------------------------------------------------------------
            // 3. Build the report.
            //    The template can reference the data source members using the
            //    syntax <<[people.Name]>> and <<[people.Age]>>.
            //    The second parameter is the data source object; the third
            //    parameter is the name used inside the template (optional).
            // -----------------------------------------------------------------
            ReportingEngine engine = new ReportingEngine
            {
                // Optional: remove empty paragraphs that may appear after
                // template tags are replaced with empty values.
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };

            // BuildReport returns a bool indicating success when InlineErrorMessages
            // option is set; we ignore the return value here.
            engine.BuildReport(templateDoc, people, "people");

            // -----------------------------------------------------------------
            // 4. Save the generated document.
            //    The Save method determines the format from the file extension.
            // -----------------------------------------------------------------
            string outputPath = @"C:\Output\PersonReport.docx";
            templateDoc.Save(outputPath);
        }
    }
}
