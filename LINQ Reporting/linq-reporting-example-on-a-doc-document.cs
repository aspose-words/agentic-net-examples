using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsLinqReportingExample
{
    // Simple data class that will be used as the LINQ data source.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a template document in memory.
            // The template uses LINQ Reporting syntax:
            //   <<[persons.Name]>> – reference a member of the data source.
            //   <<foreach [persons]>> … <</foreach>> – iterate over a collection.
            Document template = new Document();                     // create blank document
            DocumentBuilder builder = new DocumentBuilder(template); // builder for the document

            // Write a title.
            builder.Writeln("People Report");
            builder.Writeln();

            // Write a header line that will be repeated for each person.
            builder.Writeln("Name\tAge");
            builder.Writeln();

            // Insert the foreach block.
            builder.Writeln("<<foreach [persons]>>");
            builder.Writeln("<<[Name]>>\t<<[Age]>>");
            builder.Writeln("<</foreach>>");

            // 2. Prepare the data source.
            List<Person> persons = new List<Person>
            {
                new Person { Name = "Alice", Age = 30 },
                new Person { Name = "Bob",   Age = 45 },
                new Person { Name = "Carol", Age = 27 }
            };

            // 3. Build the report using the ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            // The data source name "persons" must match the name used in the template tags.
            engine.BuildReport(template, persons, "persons");

            // 4. Save the populated document to disk.
            template.Save("PeopleReport.docx"); // save as DOCX (extension determines format)
        }
    }
}
