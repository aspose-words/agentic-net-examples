using System;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Replacing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a template placeholder that accesses a member of the data source object.
        // The ":markdown" suffix tells the LINQ Reporting Engine to treat the value as Markdown.
        builder.Writeln("<<[person.Bio]:markdown>>");

        // Prepare the data source.
        var person = new Person
        {
            Name = "John Doe",
            // Sample markdown text – it will be rendered as formatted text in the output document.
            Bio = "# Biography\r\nJohn is a **software engineer** with 10+ years of experience.\r\n- C#\r\n- .NET\r\n- ASP.NET Core"
        };

        // Initialise the reporting engine.
        ReportingEngine engine = new ReportingEngine();

        // Build the report. The second argument is the data source object,
        // the third argument is the name used to reference the object inside the template.
        engine.BuildReport(doc, person, "person");

        // Save the generated report.
        doc.Save("Report.docx");
    }

    // Simple POCO used as the data source for the template.
    public class Person
    {
        public string Name { get; set; }
        public string Bio { get; set; }
    }
}
