using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsReportingDemo
{
    // Simple data model with nested members.
    public class Person
    {
        public string Name { get; set; }
        public Address Address { get; set; }
    }

    public class Address
    {
        public string City { get; set; }
        public string Street { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Prepare sample data.
            var people = new List<Person>
            {
                new Person
                {
                    Name = "John Doe",
                    Address = new Address { City = "New York", Street = "5th Avenue" }
                },
                new Person
                {
                    Name = "Jane Smith",
                    Address = new Address { City = "London", Street = "Baker Street" }
                }
            };

            // Create a new blank document and write a LINQ Reporting template.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // The template uses DOT notation to access nested members.
            // <<foreach [in persons]>> iterates over the collection.
            // Inside the loop we can reference members directly: <<[Name]>> and <<[Address.City]>>.
            builder.Writeln("<<foreach [in persons]>>");
            builder.Writeln("Name: <<[Name]>>");
            builder.Writeln("City: <<[Address.City]>>");
            builder.Writeln("<</foreach>>");

            // Build the report. The data source name "persons" matches the name used in the template.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, new object[] { people }, new[] { "persons" });

            // Save the generated report.
            doc.Save("ReportingEngine_DotNotation.docx");
        }
    }
}
