using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsXpsExport
{
    // Simple data model that will be used as the data source for the template.
    public class Person
    {
        public string Name { get; set; }
        public string Title { get; set; }
        public int Age { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Path to the DOCX template that contains placeholders like <<[person.Name]>>.
            const string templatePath = @"C:\Templates\PeopleReport.docx";

            // Load the template document (lifecycle rule: load).
            Document doc = new Document(templatePath);

            // Prepare a collection of data objects that will be merged into the template.
            List<Person> people = new List<Person>
            {
                new Person { Name = "Alice Johnson", Title = "Project Manager", Age = 34 },
                new Person { Name = "Bob Smith", Title = "Developer", Age = 28 },
                new Person { Name = "Carol White", Title = "Designer", Age = 31 }
            };

            // Use the ReportingEngine to populate the template with the data source.
            // The data source name "person" matches the placeholder prefix in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, people, "person");

            // Configure XPS save options (lifecycle rule: create).
            XpsSaveOptions xpsOptions = new XpsSaveOptions
            {
                // Example: enable high‑quality rendering for better visual fidelity.
                UseHighQualityRendering = true,
                // Example: embed the generator name (default true) – can be left unchanged.
                ExportGeneratorName = true
            };

            // Save the populated document as XPS (lifecycle rule: save).
            const string outputPath = @"C:\Output\PeopleReport.xps";
            doc.Save(outputPath, xpsOptions);

            Console.WriteLine("Document has been successfully exported to XPS.");
        }
    }
}
