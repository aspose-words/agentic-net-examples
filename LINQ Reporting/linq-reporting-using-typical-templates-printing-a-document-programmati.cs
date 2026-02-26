using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsReportingExample
{
    // Sample data class that will be used as a data source for the report.
    public class Person
    {
        public string FirstName { get; set; }
        public string LastName  { get; set; }
        public int    Age       { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // 1. Load a DOTX template that contains LINQ Reporting tags.
            //    The template file must exist at the specified path.
            Document template = new Document(@"C:\Templates\ReportTemplate.dotx");

            // 2. Prepare the data source.
            //    This can be any non‑dynamic, non‑anonymous .NET type.
            List<Person> people = new List<Person>
            {
                new Person { FirstName = "John",  LastName = "Doe",   Age = 30 },
                new Person { FirstName = "Jane",  LastName = "Smith", Age = 25 },
                new Person { FirstName = "Alice", LastName = "Brown", Age = 28 }
            };

            // 3. Create the ReportingEngine and build the report.
            //    The data source name ("persons") must match the name used in the template tags.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(template, people, "persons");

            // 4. Save the generated document.
            //    The format is inferred from the file extension.
            template.Save(@"C:\Output\GeneratedReport.docx");
        }
    }
}
