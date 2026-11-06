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
        // Create the template document with a LINQ Reporting expression that concatenates first and last name.
        const string templatePath = "Template.docx";
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Full Name: <<[person.FirstName + \" \" + person.LastName]>>");
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        var reportDoc = new Document(templatePath);

        // Prepare sample data.
        var person = new Person
        {
            FirstName = "John",
            LastName = "Doe"
        };

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, person, "person");

        // Save the generated report.
        const string outputPath = "Report.docx";
        reportDoc.Save(outputPath);
    }
}
