using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExample
{
    // Simple data class that will be used as a data source for the report.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public decimal Salary { get; set; }
    }

    public class Program
    {
        public static void Main()
        {
            // 1. Prepare the data source – a list of Person objects.
            List<Person> people = new List<Person>
            {
                new Person { Name = "Alice", Age = 30, Salary = 72000m },
                new Person { Name = "Bob",   Age = 45, Salary = 95000m },
                new Person { Name = "Carol", Age = 27, Salary = 58000m }
            };

            // 2. Load the DOCX template that contains LINQ Reporting tags,
            //    e.g. <<foreach [people]>><<[Name]>> (<<[Age]>> years) <<</foreach>>.
            Document template = new Document("Template.docx");

            // 3. Build the report using the ReportingEngine.
            //    The third argument ("people") is the name used inside the template to reference the data source.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(template, people, "people");

            // 4. Save the populated document to a new DOCX file.
            template.Save("Report.docx");
        }
    }
}
