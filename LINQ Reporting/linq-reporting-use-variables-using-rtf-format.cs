using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert template text that contains LINQ Reporting placeholders.
        // The placeholders reference members of the data source named "Data".
        builder.Writeln("Name: <<[Data.Name]>>");
        builder.Writeln("Age: <<[Data.Age]>>");

        // Insert a placeholder for a document variable.
        builder.Writeln("Custom variable: <<[MyVar]>>");

        // Add a document variable that can be accessed directly in the template.
        doc.Variables.Add("MyVar", "Variable value");

        // Prepare a simple data source object.
        var person = new Person
        {
            Name = "John Doe",
            Age = 30
        };

        // Build the report using the ReportingEngine.
        // The third argument is the name used in the template to reference the data source.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, person, "Data");

        // Configure RTF save options (e.g., compact size, compatibility with old readers).
        RtfSaveOptions rtfOptions = new RtfSaveOptions
        {
            ExportCompactSize = true,
            ExportImagesForOldReaders = false
        };

        // Save the populated document as an RTF file.
        doc.Save("Report.rtf", rtfOptions);
    }

    // Simple POCO class used as the data source for the report.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }
}
