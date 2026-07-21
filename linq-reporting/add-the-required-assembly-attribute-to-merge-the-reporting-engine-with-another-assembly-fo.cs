using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    // Public property accessed by the template.
    public string Name { get; set; } = string.Empty;
}

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "Template.docx";
        const string reportPath = "Report.docx";

        // -----------------------------------------------------------------
        // Create a simple template document containing a LINQ Reporting tag.
        // -----------------------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Customer Name: <<[person.Name]>>");
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // Load the template document for report generation.
        // -------------------------------------------------
        var reportDoc = new Document(templatePath);

        // -------------------------------------------------
        // Prepare the data model that matches the template.
        // -------------------------------------------------
        var person = new Person { Name = "John Doe" };

        // -------------------------------------------------
        // Build the report using the ReportingEngine.
        // -------------------------------------------------
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, person, "person");

        // -------------------------------------------------
        // Save the generated report.
        // -------------------------------------------------
        reportDoc.Save(reportPath);
    }
}
