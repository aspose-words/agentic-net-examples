using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare a hierarchical data source.
        var person = new Person
        {
            Name = "John Doe",
            Age = 30,
            Address = new Address
            {
                City = "New York",
                Street = "5th Avenue",
                Zip = "10001"
            }
        };

        // Create an in‑memory DOTM template.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use contextual member access (DOTM format) in the template.
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Age: <<[person.Age]>>");
        builder.Writeln("City: <<[person.Address.City]>>");
        builder.Writeln("Full address: <<[person.Address.Street]>>, <<[person.Address.City]>> <<[person.Address.Zip]>>");

        // Build the report, exposing the root object as "person".
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.AllowMissingMembers
        };
        engine.BuildReport(doc, person, "person");

        // Save the result as a macro‑enabled template (DOTM).
        doc.Save("Report.dotm");
    }

    // Sample data classes.
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
        public string Zip { get; set; }
    }
}
