using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingMhtmlExample
{
    // Simple data source class used by the LINQ Reporting Engine.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a new blank Word document (lifecycle: create).
            Document doc = new Document();

            // 2. Build a template using DocumentBuilder.
            //    The template uses LINQ Reporting Engine syntax: <<[person.Name]>> and <<[person.Age]>>.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Person Report");
            builder.Writeln("Name: <<[person.Name]>>");
            builder.Writeln("Age: <<[person.Age]>>");

            // 3. Prepare the data source.
            Person person = new Person { Name = "John Doe", Age = 42 };

            // 4. Create the ReportingEngine and populate the template (LINQ Reporting Engine API).
            ReportingEngine engine = new ReportingEngine();
            // The third argument is the name used to reference the data source inside the template.
            engine.BuildReport(doc, person, "person");

            // 5. Save the populated document as MHTML (lifecycle: save).
            //    SaveFormat.Mhtml ensures the correct output format.
            doc.Save("PersonReport.mhtml", SaveFormat.Mhtml);
        }
    }
}
