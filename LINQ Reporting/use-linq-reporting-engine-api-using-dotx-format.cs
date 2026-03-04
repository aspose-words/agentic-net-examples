using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Path to the DOTX template file.
        string templatePath = "Template.dotx";

        // Path where the generated report will be saved.
        string outputPath = "Report.dotx";

        // Load the DOTX template document.
        Document doc = new Document(templatePath);

        // Create a simple data source object.
        var data = new Person
        {
            FirstName = "John",
            LastName = "Doe",
            Age = 30
        };

        // Initialize the LINQ Reporting Engine.
        ReportingEngine engine = new ReportingEngine();

        // Populate the template with data.
        // The data source name "person" can be referenced in the template as <<[person.FirstName]>> etc.
        engine.BuildReport(doc, data, "person");

        // Save the populated document as a DOTX file.
        doc.Save(outputPath, SaveFormat.Dotx);
    }

    // Simple POCO class used as a data source for the template.
    public class Person
    {
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public int Age { get; set; }
    }
}
