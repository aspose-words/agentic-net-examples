using System;
using System.IO;
using System.Text;
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
        // Register code page provider for Aspose.Words if needed.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare folders.
        string outputDir = "output";
        Directory.CreateDirectory(outputDir);

        // Create a template document with a LINQ Reporting expression that concatenates first and last name.
        string templatePath = Path.Combine(outputDir, "template.docx");
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Full Name: <<[person.FirstName + \" \" + person.LastName]>>");
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        var doc = new Document(templatePath);

        // Create sample data.
        var person = new Person
        {
            FirstName = "John",
            LastName = "Doe"
        };

        // Build the report.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, person, "person");

        // Save the generated report.
        string resultPath = Path.Combine(outputDir, "result.docx");
        doc.Save(resultPath);

        // Indicate completion.
        Console.WriteLine($"Report generated: {resultPath}");
    }
}
