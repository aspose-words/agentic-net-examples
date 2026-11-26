using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data model.
        var model = new ReportModel
        {
            Person = new Person { Name = "John Doe" }
        };

        // Create a template document programmatically.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Hello <<[model.Person.Name]>>!");
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        var reportDoc = new Document(templatePath);

        // Configure the ReportingEngine.
        var engine = new ReportingEngine();

        // NOTE: In newer versions of Aspose.Words the RestrictedMembers property may not be available.
        // If it is present, you could add disallowed members here to block file‑writing methods.
        // Example (when supported):
        // engine.RestrictedMembers = new RestrictedMembers();
        // engine.RestrictedMembers.AddDisallowedMember("System.IO.File", "WriteAllText");
        // engine.RestrictedMembers.AddDisallowedMember("System.IO.File", "WriteAllBytes");
        // engine.RestrictedMembers.AddDisallowedMember("System.IO.File", "WriteAllLines");
        // engine.RestrictedMembers.AddDisallowedMember("System.IO.StreamWriter", "Write");
        // engine.RestrictedMembers.AddDisallowedMember("System.IO.StreamWriter", "WriteLine");

        // Build the report.
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");
        reportDoc.Save(outputPath);

        Console.WriteLine($"Report generated at: {outputPath}");
    }
}

// Data model classes.
public class ReportModel
{
    public Person Person { get; set; } = new();
}

public class Person
{
    public string Name { get; set; } = string.Empty;
}
