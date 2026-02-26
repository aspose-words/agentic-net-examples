using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace AsposeWordsLinqReportingToXps
{
    // Simple data model used as the reporting data source.
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
            // Load the PDF template that contains LINQ Reporting tags (e.g. <<[ds.Name]>>).
            Document document = new Document("Template.pdf");

            // Prepare a collection of data objects to populate the template.
            List<Person> data = new List<Person>
            {
                new Person { Name = "Alice", Age = 30, City = "New York" },
                new Person { Name = "Bob",   Age = 45, City = "London"   },
                new Person { Name = "Carol", Age = 27, City = "Tokyo"    }
            };

            // Create the reporting engine and build the report.
            ReportingEngine engine = new ReportingEngine();
            // The data source name ("ds") must match the name used in the template tags.
            engine.BuildReport(document, data, "ds");

            // Save the populated document as XPS using XpsSaveOptions.
            XpsSaveOptions xpsOptions = new XpsSaveOptions();
            document.Save("ReportOutput.xps", xpsOptions);
        }
    }
}
