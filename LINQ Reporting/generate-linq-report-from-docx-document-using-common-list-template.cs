using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportExample
{
    // Simple data class that will be used as the data source for the report.
    public class Person
    {
        public string FirstName { get; set; }
        public string LastName  { get; set; }
        public int    Age       { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // Load the DOCX template that contains the Common List syntax (e.g. <<foreach [in persons]>><<[FirstName]>><</foreach>>).
            Document template = new Document("Template.docx");

            // Prepare a collection of objects that will be bound to the template.
            List<Person> persons = new List<Person>
            {
                new Person { FirstName = "John",  LastName = "Doe",   Age = 30 },
                new Person { FirstName = "Jane",  LastName = "Smith", Age = 25 },
                new Person { FirstName = "Bob",   LastName = "Brown", Age = 40 }
            };

            // Create the reporting engine and populate the template with the data source.
            ReportingEngine engine = new ReportingEngine();
            // The third argument is the name used inside the template to reference the data source.
            engine.BuildReport(template, persons, "persons");

            // Save the generated report.
            template.Save("Report.docx");
        }
    }
}
