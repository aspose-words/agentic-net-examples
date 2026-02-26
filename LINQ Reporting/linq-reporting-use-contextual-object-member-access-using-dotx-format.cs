using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

namespace AsposeWordsDotxExample
{
    // Simple data model with nested objects to demonstrate DOTX member access.
    public class Person
    {
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public Address HomeAddress { get; set; }
    }

    public class Address
    {
        public string Street { get; set; }
        public string City { get; set; }
        public string Country { get; set; }
    }

    class Program
    {
        static void Main()
        {
            // 1. Create a new blank document.
            Document doc = new Document();

            // 2. Build a template using DocumentBuilder.
            //    The template uses DOTX syntax (dot notation) to access nested members.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("<<[person.FirstName]>> <<[person.LastName]>>");
            builder.Writeln("Address: <<[person.HomeAddress.Street]>>, <<[person.HomeAddress.City]>>, <<[person.HomeAddress.Country]>>");

            // 3. Prepare the data source.
            Person person = new Person
            {
                FirstName = "John",
                LastName = "Doe",
                HomeAddress = new Address
                {
                    Street = "123 Main St",
                    City = "Springfield",
                    Country = "USA"
                }
            };

            // 4. Build the report using ReportingEngine.
            //    The third parameter is the name used in the template to reference the data source object.
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(doc, person, "person");

            // 5. Save the resulting document.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "ReportWithDotx.docx");
            doc.Save(outputPath);
        }
    }
}
