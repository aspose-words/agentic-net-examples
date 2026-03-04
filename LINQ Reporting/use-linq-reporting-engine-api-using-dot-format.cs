using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsReportingEngineExample
{
    // Simple data class used as the data source for the report.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Create a new blank Word document.
            Document doc = new Document();

            // Use DocumentBuilder to insert template tags that the LINQ Reporting Engine will replace.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Name: <<[person.Name]>>");
            builder.Writeln("Age: <<[person.Age]>>");

            // Prepare the data source object.
            Person person = new Person
            {
                Name = "John Doe",
                Age = 30
            };

            // Initialize the ReportingEngine and build the report.
            ReportingEngine engine = new ReportingEngine();
            // The third parameter ("person") is the name used to reference the data source in the template.
            engine.BuildReport(doc, person, "person");

            // Save the populated document as a DOT (Word template) file.
            doc.Save("ReportTemplate.dot", SaveFormat.Dot);
        }
    }
}
