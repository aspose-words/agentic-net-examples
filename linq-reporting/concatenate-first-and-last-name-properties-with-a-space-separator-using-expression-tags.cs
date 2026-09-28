using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    public string FirstName { get; set; } = "";
    public string LastName { get; set; } = "";
}

public class Program
{
    public static void Main()
    {
        // Paths for template and output documents.
        const string templatePath = "template.docx";
        const string outputPath = "output.docx";

        // Create a template document with a LINQ Reporting expression that concatenates first and last name.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Full Name: <<[person.FirstName + \" \" + person.LastName]>>");
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        var doc = new Document(templatePath);

        // Sample data.
        var person = new Person
        {
            FirstName = "John",
            LastName = "Doe"
        };

        // Build the report using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, person, "person");

        // Save the generated report.
        doc.Save(outputPath);
    }
}
