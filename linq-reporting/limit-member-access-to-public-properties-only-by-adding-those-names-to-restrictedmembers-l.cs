using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for any required encodings.
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Paths for the template and the generated report.
        const string templatePath = "Template.docx";
        const string reportPath = "Report.docx";

        // -------------------------------------------------
        // Create a simple template document with LINQ tags.
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Secret: <<[person.Secret]>>");
        templateDoc.Save(templatePath);

        // Load the template back for reporting.
        Document reportDoc = new Document(templatePath);

        // -------------------------------------------------
        // Prepare the data model.
        // -------------------------------------------------
        Person sourcePerson = new Person
        {
            Name = "John Doe",
            Secret = "TopSecret"
        };

        // Wrap the source object in a model that exposes only the allowed members.
        PersonReport person = new PersonReport
        {
            Name = sourcePerson.Name,
            // Secret is intentionally omitted.
        };

        // -------------------------------------------------
        // Configure the ReportingEngine.
        // -------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        // Allow missing members so that restricted members are rendered as empty strings.
        engine.Options = ReportBuildOptions.AllowMissingMembers;

        // Build the report using the root object name "person".
        engine.BuildReport(reportDoc, person, "person");

        // Save the generated report.
        reportDoc.Save(reportPath);
    }
}

// -----------------------------------------------------
// Data model with public properties.
// -----------------------------------------------------
public class Person
{
    public string Name { get; set; }
    public string Secret { get; set; }
}

// -----------------------------------------------------
// Wrapper model exposing only the allowed public properties.
// -----------------------------------------------------
public class PersonReport
{
    public string Name { get; set; }
}
