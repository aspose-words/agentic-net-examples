using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

namespace LinqReportingExample
{
    // Simple data class used in the LINQ reporting template.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a blank Word document.
            Document doc = new Document();

            // 2. Insert LINQ reporting tags into the document.
            //    The tags will be processed by ReportingEngine.BuildReport.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("People Report");
            builder.Writeln("----------------");
            builder.Writeln("Name: <<[data.People.Name]>>");
            builder.Writeln("Age : <<[data.People.Age]>>");
            builder.Writeln();
            builder.Writeln("Average Age: <<[data.People.Average(p => p.Age):0.00]>>");
            builder.Writeln("Total Age  : <<[data.People.Sum(p => p.Age)]>>");

            // 3. Prepare the data source – a list of Person objects.
            List<Person> people = new List<Person>
            {
                new Person { Name = "Alice", Age = 30 },
                new Person { Name = "Bob",   Age = 45 },
                new Person { Name = "Carol", Age = 27 }
            };

            // 4. Build the report using the LINQ Reporting Engine.
            //    The anonymous object provides a name ("data") that can be referenced in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, new { People = people }, "data");

            // 5. Save the populated document as PDF.
            doc.Save("PeopleReport.pdf", SaveFormat.Pdf);
        }
    }
}
