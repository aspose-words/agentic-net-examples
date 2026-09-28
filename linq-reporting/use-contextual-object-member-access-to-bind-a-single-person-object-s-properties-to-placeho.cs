using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    public string Name { get; set; } = "John Doe";
    public int Age { get; set; } = 30;
    public string Email { get; set; } = "john.doe@example.com";
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for template and output documents.
        string templatePath = "Template.docx";
        string outputPath = "Report.docx";

        // Create a template document with LINQ Reporting tags.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Age: <<[person.Age]>>");
        builder.Writeln("Email: <<[person.Email]>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // Load the template for report generation.
        Document reportDoc = new Document(templatePath);

        // Sample data.
        Person person = new Person
        {
            Name = "Alice Smith",
            Age = 28,
            Email = "alice.smith@example.com"
        };

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, person, "person");

        // Save the generated report.
        reportDoc.Save(outputPath);
    }
}
