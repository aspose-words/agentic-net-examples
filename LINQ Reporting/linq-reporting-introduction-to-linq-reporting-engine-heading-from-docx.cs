using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the DOCX template that contains LINQ Reporting Engine tags.
        // Example tag in the template: <<foreach [persons]>><<[Name]>> (Age: <<[Age]>>)<</foreach>>
        Document template = new Document("Template.docx");

        // Prepare a data source that will be referenced from the template.
        // Any non‑dynamic, non‑anonymous .NET type can be used; here we use a List<Person>.
        List<Person> persons = new List<Person>
        {
            new Person { Name = "John Doe", Age = 30 },
            new Person { Name = "Jane Smith", Age = 25 },
            new Person { Name = "Bob Johnson", Age = 45 }
        };

        // Create an instance of the ReportingEngine.
        ReportingEngine engine = new ReportingEngine();

        // Build the report. The third argument is the name used in the template to refer to the data source.
        engine.BuildReport(template, persons, "persons");

        // Save the populated document.
        template.Save("Report.docx");
    }

    // Simple POCO class used as the data model for the report.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }
}
