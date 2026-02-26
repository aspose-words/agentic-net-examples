using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading; // Added for LoadOptions
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a plain‑text template that uses LINQ Reporting syntax.
        // -----------------------------------------------------------------
        // The template is saved as a .txt file so that Aspose.Words loads it
        // with LoadFormat.Text. The placeholders reference members of the
        // object that will be supplied to the ReportingEngine.
        string templateContent =
            "Name: <<[person.Name]>>\n" +
            "Age:  <<[person.Age]>>\n" +
            "City: <<[person.Address.City]>>";

        const string templatePath = "Template.txt";
        File.WriteAllText(templatePath, templateContent);

        // ---------------------------------------------------------------
        // 2. Load the TXT template into a Document object.
        // ---------------------------------------------------------------
        // LoadOptions tells Aspose.Words to treat the file as plain text.
        LoadOptions loadOptions = new LoadOptions { LoadFormat = LoadFormat.Text }; // Use LoadFormat.Text
        Document doc = new Document(templatePath, loadOptions);

        // ---------------------------------------------------------------
        // 3. Prepare the data source – a simple POCO hierarchy.
        // ---------------------------------------------------------------
        var person = new Person
        {
            Name = "John Doe",
            Age = 30,
            Address = new Address
            {
                City = "London",
                Street = "Baker St"
            }
        };

        // ---------------------------------------------------------------
        // 4. Build the report. The third argument ("person") is the name
        //    used inside the template to reference the root object.
        // ---------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, person, "person");

        // ---------------------------------------------------------------
        // 5. Save the populated document.
        // ---------------------------------------------------------------
        doc.Save("Report.docx");
    }

    // -----------------------------------------------------------------
    // Simple POCO classes used as the data source.
    // -----------------------------------------------------------------
    public class Person
    {
        public string Name { get; set; } = string.Empty;
        public int Age { get; set; }
        public Address Address { get; set; } = null!;
    }

    public class Address
    {
        public string City { get; set; } = string.Empty;
        public string Street { get; set; } = string.Empty;
    }
}
