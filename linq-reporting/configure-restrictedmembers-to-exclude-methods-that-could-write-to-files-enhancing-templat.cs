using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Restrict types that expose file‑writing capabilities.
        // This must be done before any report is built.
        ReportingEngine.SetRestrictedTypes(
            typeof(System.IO.File),
            typeof(System.IO.StreamWriter),
            typeof(System.IO.FileInfo));

        // Create a simple template document with a LINQ Reporting tag.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello <<[person.Name]>>!");

        // Prepare the data model.
        var person = new Person { Name = "John Doe" };

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // default options
        engine.BuildReport(doc, person, "person");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}

// Simple data model used by the template.
public class Person
{
    public string Name { get; set; } = string.Empty;
}
