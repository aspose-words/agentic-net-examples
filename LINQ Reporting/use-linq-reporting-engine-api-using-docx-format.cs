using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsReportingDemo
{
    // Simple data entity used as a data source for the report.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Path to the DOCX template that contains Reporting Engine tags,
            // e.g. <<foreach [in persons]>><<[Name]>> (<<[Age]>>)<</foreach>>
            string templatePath = "Template.docx";

            // Load the template document.
            Document doc = new Document(templatePath);

            // Prepare a collection of data objects.
            List<Person> persons = new List<Person>
            {
                new Person { Name = "John Doe", Age = 30 },
                new Person { Name = "Jane Smith", Age = 25 },
                new Person { Name = "Bob Johnson", Age = 45 }
            };

            // Create and configure the ReportingEngine.
            ReportingEngine engine = new ReportingEngine
            {
                // Example option: remove paragraphs that become empty after processing.
                Options = ReportBuildOptions.RemoveEmptyParagraphs
            };

            // Build the report. The third argument is the name used in the template to reference the data source.
            engine.BuildReport(doc, persons, "persons");

            // Save the populated document.
            doc.Save("ReportOutput.docx");
        }
    }
}
