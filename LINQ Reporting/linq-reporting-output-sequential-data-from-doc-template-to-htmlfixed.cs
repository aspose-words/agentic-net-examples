using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the DOC template that contains LINQ Reporting tags.
        Document template = new Document("Template.docx");

        // Prepare a simple data source – a list of Person objects.
        var people = new List<Person>
        {
            new Person { Name = "John Doe", Age = 30 },
            new Person { Name = "Jane Smith", Age = 25 },
            new Person { Name = "Bob Johnson", Age = 42 }
        };

        // Create the reporting engine and populate the template.
        ReportingEngine engine = new ReportingEngine();
        // The third argument is the name used to reference the data source inside the template.
        engine.BuildReport(template, people, "people");

        // Configure HTML Fixed save options.
        HtmlFixedSaveOptions htmlOptions = new HtmlFixedSaveOptions
        {
            SaveFormat = SaveFormat.HtmlFixed,   // Ensure the format is HTML Fixed.
            ExportEmbeddedImages = true,        // Embed images directly into the HTML.
            ShowPageBorder = false              // Optional: hide page borders in the output.
        };

        // Save the populated document as HTML Fixed.
        template.Save("Report.html", htmlOptions);
    }

    // Simple POCO class used as the data source.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }
}
