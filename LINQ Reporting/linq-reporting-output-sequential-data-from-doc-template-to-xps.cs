using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the DOCX template that contains LINQ Reporting tags,
        // e.g. <<foreach [people]>> <<[Name]>> <<[Age]>> <</foreach>>
        Document template = new Document("Template.docx");

        // Prepare a sequential data source.
        List<Person> people = new List<Person>
        {
            new Person { Name = "John Doe",   Age = 30 },
            new Person { Name = "Jane Smith", Age = 25 },
            new Person { Name = "Bob Johnson",Age = 40 }
        };

        // Create the reporting engine and populate the template.
        ReportingEngine engine = new ReportingEngine();
        // The third argument is the name used inside the template tags.
        engine.BuildReport(template, people, "people");

        // Save the populated document as XPS.
        XpsSaveOptions xpsOptions = new XpsSaveOptions();
        template.Save("Report.xps", xpsOptions);
    }

    // Simple data class used as the data source.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }
}
