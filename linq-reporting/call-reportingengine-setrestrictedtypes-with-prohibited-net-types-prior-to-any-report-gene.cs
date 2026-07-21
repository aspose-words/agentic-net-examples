using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create a simple data model.
        var person = new Person { Name = "John Doe" };

        // Build a template document programmatically.
        var templatePath = "Template.docx";
        var builder = new DocumentBuilder();
        // Insert a LINQ Reporting tag that references the data model.
        builder.Writeln("Customer name: <<[person.Name]>>");
        // Save the template to disk.
        builder.Document.Save(templatePath);

        // Load the template document.
        var doc = new Document(templatePath);

        // Specify prohibited .NET types before any report generation.
        ReportingEngine.SetRestrictedTypes(typeof(System.Environment), typeof(System.IO.File));

        // Build the report using the template and the data model.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, person, "person");

        // Save the generated report.
        doc.Save("Report.docx");
    }
}

// Simple data model with a non‑nullable property to avoid warnings.
public class Person
{
    public string Name { get; set; } = string.Empty;
}
