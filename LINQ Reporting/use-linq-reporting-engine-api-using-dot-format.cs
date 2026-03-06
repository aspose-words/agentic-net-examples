using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Path to the DOT template file.
        string templatePath = "Template.dot";

        // Path where the generated report will be saved.
        string outputPath = "Report.docx";

        // Ensure a simple DOT template exists.
        CreateTemplateIfMissing(templatePath);

        // Load the DOT template into a Document object.
        Document template = new Document(templatePath);

        // Prepare a data source object. Its members will be referenced in the template.
        var person = new Person
        {
            Name = "John Doe",
            Age = 30
        };

        // Initialize the LINQ Reporting Engine.
        ReportingEngine engine = new ReportingEngine();

        // Build the report by populating the template with the data source.
        // The third argument ("person") is the name used in the template tags.
        engine.BuildReport(template, person, "person");

        // Save the populated document as a regular DOCX file.
        template.Save(outputPath, SaveFormat.Docx);
    }

    // Creates a minimal DOT template containing LINQ tags if it does not already exist.
    static void CreateTemplateIfMissing(string path)
    {
        if (File.Exists(path))
            return;

        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // LINQ tags that reference the data source members.
        builder.Writeln("Report for <<[person.Name]>>");
        builder.Writeln("Age: <<[person.Age]>>");

        // Save the document as a DOT (Word template) file.
        doc.Save(path, SaveFormat.Dot);
    }

    // Simple POCO class used as the data source for the report.
    public class Person
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }
}
