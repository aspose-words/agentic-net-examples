using System;
using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a heading for the report.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("LINQ Reporting Introduction to LINQ Reporting Engine");

        // Insert a placeholder that will be replaced by the data source.
        builder.Writeln("Name: <<[person.Name]>>");

        // Prepare a simple data source object.
        var person = new Person { Name = "John Doe" };

        // Use the LINQ Reporting Engine to populate the template.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, person, "person");

        // Save the resulting document.
        doc.Save("LINQReportingExample.docx");
    }
}

// Simple POCO class used as a data source.
public class Person
{
    public string Name { get; set; }
}
