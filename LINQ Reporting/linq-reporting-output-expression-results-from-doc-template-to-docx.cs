using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Load the DOC template that contains LINQ Reporting tags, e.g. <<foreach [people]>><<[FirstName]>><</foreach>>
        Document template = new Document("Template.doc");

        // Prepare a data source that the template will consume.
        var people = new List<Person>
        {
            new Person { FirstName = "John",  LastName = "Doe",   Age = 30 },
            new Person { FirstName = "Jane",  LastName = "Smith", Age = 25 },
            new Person { FirstName = "Alice", LastName = "Brown", Age = 28 }
        };

        // Create the reporting engine.
        ReportingEngine engine = new ReportingEngine();

        // Populate the template. The third argument is the name used inside the template to refer to the data source.
        engine.BuildReport(template, people, "people");

        // Save the generated report as DOCX.
        template.Save("Result.docx", SaveFormat.Docx);
    }

    // Simple POCO class used as a data source for the report.
    public class Person
    {
        public string FirstName { get; set; }
        public string LastName  { get; set; }
        public int Age { get; set; }
    }
}
