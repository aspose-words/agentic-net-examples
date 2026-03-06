using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReport
{
    // Simple data model that will be used as the LINQ data source.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public string City { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Path to the template document that contains Reporting Engine tags,
            // e.g. <<[ds.Name]>>, <<foreach [ds]>>...<<[Name]>>...<<[/foreach]>>.
            string templatePath = @"C:\Docs\Template.docx";

            // Load the template document.
            Document template = new Document(templatePath);

            // Prepare a LINQ data source – a list of Person objects.
            List<Person> people = new List<Person>
            {
                new Person { Name = "Alice", Age = 30, City = "New York" },
                new Person { Name = "Bob",   Age = 45, City = "London"   },
                new Person { Name = "Carol", Age = 27, City = "Tokyo"    }
            };

            // The ReportingEngine can work with any non‑dynamic .NET object.
            // Here we pass the list directly; the engine will treat it as a collection.
            ReportingEngine engine = new ReportingEngine();

            // Build the report. The data source name ("ds") must match the name used in the template.
            engine.BuildReport(template, people, "ds");

            // Save the populated report.
            string outputPath = @"C:\Docs\ReportResult.docx";
            template.Save(outputPath);
        }
    }
}
