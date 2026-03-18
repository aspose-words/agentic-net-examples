using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsReportingDemo
{
    // External type whose properties will be used inside the template.
    public class Person
    {
        public required string Name { get; set; }
        public int Age { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Prepare a collection of Person objects.
            var people = new List<Person>
            {
                new Person { Name = "Alice",   Age = 28 },
                new Person { Name = "Bob",     Age = 35 },
                new Person { Name = "Charlie", Age = 42 }
            };

            // Create the reporting engine and register the external type.
            ReportingEngine engine = new ReportingEngine();
            engine.KnownTypes.Add(typeof(Person));

            // Create a simple template document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a title.
            builder.Writeln("Adults (Age > 30):");

            // Use the Where extension method with a lambda that references Person.Age.
            // The result is stored in a temporary variable "adults".
            builder.InsertField("<<var adults = Model.Where(p => p.Age > 30)>>");

            // Iterate over the filtered collection and output each person's name.
            builder.InsertField("<<foreach var a in adults>>");
            builder.Writeln(" - <<a.Name>>");
            builder.InsertField("<</foreach>>");

            // Build the report using the ReportingEngine.
            engine.BuildReport(doc, people);

            // Save the generated document.
            doc.Save("AdultsReport.docx");
        }
    }
}
