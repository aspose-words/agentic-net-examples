using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace LinqReportingExample
{
    // Simple data class for LINQ source
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a blank document that will serve as a template.
            Document template = new Document();

            // 2. Insert template tags using DocumentBuilder.
            DocumentBuilder builder = new DocumentBuilder(template);
            builder.Writeln("<<[people.Name]>> is <<[people.Age]>> years old.");
            builder.Writeln("<<foreach [people]>>");
            builder.Writeln("- <<[Name]>> (<<[Age]>>)");
            builder.Writeln("<</foreach>>");

            // 3. Prepare LINQ data source.
            List<Person> people = new List<Person>
            {
                new Person { Name = "Alice", Age = 28 },
                new Person { Name = "Bob",   Age = 35 },
                new Person { Name = "Carol", Age = 42 }
            };

            // 4. Build the report using ReportingEngine.
            ReportingEngine engine = new ReportingEngine();
            // The data source name "people" matches the tags used in the template.
            engine.BuildReport(template, people, "people");

            // 5. Save the populated document.
            template.Save("LinqReport.docx");
        }
    }
}
