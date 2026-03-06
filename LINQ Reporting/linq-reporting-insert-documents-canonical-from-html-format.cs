using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the HTML template that contains ReportingEngine tags, e.g. <<foreach [people]>><<[Name]>> (<[Age]>)<</foreach>>
        Document template = new Document("Template.html");

        // Prepare a data source that can be referenced from the template.
        // The data source must be a concrete .NET type (not anonymous) for ReportingEngine.
        List<Person> people = new List<Person>
        {
            new Person { Name = "John Doe", Age = 30 },
            new Person { Name = "Jane Smith", Age = 25 },
            new Person { Name = "Bob Johnson", Age = 45 }
        };

        // Create the reporting engine.
        ReportingEngine engine = new ReportingEngine();

        // Build the report. The third argument is the name used inside the template to reference the data source.
        engine.BuildReport(template, people, "people");

        // Save the populated document.
        template.Save("Report.docx");
    }

    // Simple POCO class used as the data source.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }
}
