using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Create a blank Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert template tags that access members of the data source object.
        // The data source will be referenced by the name "person".
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Age: <<[person.Age]>>");
        builder.Writeln("City: <<[person.Address.City]>>");

        // Prepare the data source object.
        var person = new Person
        {
            Name = "John Doe",
            Age = 30,
            Address = new Address
            {
                City = "New York",
                Street = "5th Avenue"
            }
        };

        // Build the report, passing the document, the data source object,
        // and the name that will be used to reference the object in the template.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, person, "person");

        // Save the populated document.
        doc.Save("Report.docx");
    }

    // Sample data classes used as the reporting data source.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public Address Address { get; set; }
    }

    public class Address
    {
        public string City { get; set; }
        public string Street { get; set; }
    }
}
