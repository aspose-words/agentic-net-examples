using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    public string Name { get; set; } = string.Empty;
}

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some Aspose.Words features).
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Create a template document programmatically.
        DocumentBuilder builder = new DocumentBuilder();
        // Normal member.
        builder.Writeln("Name: <<[person.Name]>>");
        // Missing member – will be treated as null.
        builder.Writeln("Missing: <<[person.NonExisting]>>");
        // Save the template (optional, just to have a file on disk).
        const string templatePath = "Template.docx";
        builder.Document.Save(templatePath);

        // Load the template document.
        Document doc = new Document(templatePath);

        // Prepare the data source.
        Person person = new Person { Name = "John Doe" };

        // Configure the reporting engine to treat missing members as null.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.AllowMissingMembers;
        engine.MissingMemberMessage = "N/A";

        // Build the report.
        engine.BuildReport(doc, person, "person");

        // Save the generated report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}
