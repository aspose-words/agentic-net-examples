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
            // Load the DOCX template that contains an in‑paragraph list template.
            // Example template syntax:
            //   <<foreach [people]>>
            //   <<list [people]>> <<[Name]>> (<<[Age]>> years) from <<[City]>>
            //   <</foreach>>
            Document template = new Document("TemplateWithInParagraphList.docx");

            // Prepare a LINQ data source – a list of Person objects.
            List<Person> people = new List<Person>
            {
                new Person { Name = "Alice", Age = 30, City = "New York" },
                new Person { Name = "Bob",   Age = 25, City = "London"   },
                new Person { Name = "Carol", Age = 28, City = "Paris"    }
            };

            // Create the reporting engine.
            ReportingEngine engine = new ReportingEngine();

            // Build the report by populating the template with the data source.
            // The data source name ("people") matches the name used in the template tags.
            engine.BuildReport(template, people, "people");

            // Save the generated report.
            template.Save("LinqReport_Output.docx");
        }
    }
}
