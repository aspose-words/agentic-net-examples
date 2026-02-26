using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingDemo
{
    // Simple data class that will be used as the data source for the report.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a blank Word document that will serve as the template.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Insert a heading placeholder that the LINQ Reporting Engine will replace.
            // The syntax "<<[person.Name]>>" tells the engine to insert the Name property of the data source.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Writeln("<<[person.Name]>>");

            // Insert a normal paragraph with another placeholder.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
            builder.Writeln("Age: <<[person.Age]>>");

            // 2. Prepare the data source.
            Person person = new Person { Name = "John Doe", Age = 42 };

            // 3. Create the ReportingEngine and build the report.
            ReportingEngine engine = new ReportingEngine();
            // The second parameter is the data source object.
            // The third parameter ("person") is the name used to reference the object inside the template.
            engine.BuildReport(template, person, "person");

            // 4. Save the populated document.
            template.Save("LinqReportingResult.docx");
        }
    }
}
