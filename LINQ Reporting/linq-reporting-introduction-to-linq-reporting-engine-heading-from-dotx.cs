using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingDemo
{
    // Simple data source class – any non‑dynamic, non‑anonymous type can be used.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Load the DOTX template that contains the LINQ Reporting Engine heading.
            // The constructor Document(string) loads a document from the file system.
            Document template = new Document(@"Templates\LinqReportingTemplate.dotx");

            // Prepare a data source instance.
            Person data = new Person
            {
                Name = "John Doe",
                Age = 30
            };

            // Create the reporting engine and populate the template.
            // BuildReport(Document, object, string) allows the template to reference the data source object itself.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(template, data, "person");

            // Save the generated report. The Save(string) overload determines the format from the file extension.
            template.Save(@"Results\LinqReportingResult.docx");
        }
    }
}
