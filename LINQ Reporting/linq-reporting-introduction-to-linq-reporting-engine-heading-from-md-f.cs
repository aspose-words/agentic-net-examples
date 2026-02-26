using Aspose.Words;
using Aspose.Words.Reporting;

class Program
{
    static void Main()
    {
        // Create a new document and add a heading for the LINQ Reporting Engine introduction
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("LINQ Reporting Introduction to LINQ Reporting Engine");

        // Insert a placeholder that will be replaced by the data source during report generation
        builder.Writeln("Name: <<[person.Name]>>");

        // Prepare a simple data source object
        var person = new Person { Name = "John Doe" };

        // Use ReportingEngine to populate the template with the data source
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, person, "person");

        // Save the generated report
        doc.Save("Report.docx");
    }
}

// Simple data class used as the data source for the report
public class Person
{
    public string Name { get; set; }
}
