using System;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsReportingExample
{
    // Simple data model used as the data source for the report.
    public class Address
    {
        public string City { get; set; }
        public string Street { get; set; }
    }

    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public Address Address { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // Create a blank document that will serve as the template.
            Document template = new Document();
            DocumentBuilder builder = new DocumentBuilder(template);

            // Insert WORDML (LINQ Reporting) tags that access members of the data source.
            // The data source will be named "person", so we reference its members using that name.
            builder.Writeln("Name: <<[person.Name]>>");
            builder.Writeln("Age: <<[person.Age]>>");
            builder.Writeln("City: <<[person.Address.City]>>");
            builder.Writeln("Street: <<[person.Address.Street]>>");

            // Prepare the data source object.
            Person person = new Person
            {
                Name = "John Doe",
                Age = 42,
                Address = new Address
                {
                    City = "New York",
                    Street = "5th Avenue"
                }
            };

            // Configure the reporting engine.
            ReportingEngine engine = new ReportingEngine
            {
                // Allow missing members to be ignored (optional).
                Options = ReportBuildOptions.AllowMissingMembers
            };

            // Build the report. The second parameter is the data source object,
            // the third parameter is the name used inside the template to reference it.
            engine.BuildReport(template, person, "person");

            // Save the populated document.
            template.Save("Report_Output.docx");
        }
    }
}
