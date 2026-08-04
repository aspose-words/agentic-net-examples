using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    public string Name { get; set; } = "";
    public int Age { get; set; }
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for legacy encodings.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        const string templatePath = "template.docx";
        const string outputPath = "report.docx";

        // -------------------------------------------------
        // Create the template document programmatically.
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Use default LINQ Reporting tags << >> to avoid tag collisions.
        builder.Writeln("<<[person.Name]>> is <<[person.Age]>> years old.");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // Load the template back for reporting.
        // -------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // Sample data source.
        Person person = new Person
        {
            Name = "John Doe",
            Age = 30
        };

        // -------------------------------------------------
        // Configure the ReportingEngine.
        // -------------------------------------------------
        ReportingEngine engine = new ReportingEngine();

        // Build the report using the root object name "person".
        engine.BuildReport(reportDoc, person, "person");

        // Save the generated report.
        reportDoc.Save(outputPath);
    }
}
