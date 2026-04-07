using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider required by Aspose.Words for some encodings.
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Prepare sample data.
        var person = new Person
        {
            Name = "John Doe",
            Age = 30
        };

        // Create a simple template document with LINQ Reporting tags.
        const string templatePath = "Template.docx";
        CreateTemplate(templatePath);

        // Load the template before building the report.
        var doc = new Document(templatePath);

        // Restrict access to the Person type (all its members become inaccessible).
        // This must be done before the first report build.
        ReportingEngine.SetRestrictedTypes(typeof(Person));

        // Configure the reporting engine.
        var engine = new ReportingEngine
        {
            // Allow missing members so that restricted members are treated as null
            // instead of throwing an exception.
            Options = ReportBuildOptions.AllowMissingMembers,
            MissingMemberMessage = string.Empty
        };

        // Build the report. The root object name must match the tag prefix used in the template.
        engine.BuildReport(doc, person, "person");

        // Save the generated report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);

        Console.WriteLine($"Report generated successfully: {Path.GetFullPath(outputPath)}");
    }

    // Helper method to create the template document.
    private static void CreateTemplate(string filePath)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // The template contains two tags: one for Name and one for Age.
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Age: <<[person.Age]>>");

        doc.Save(filePath);
    }
}

// Simple data model used by the report.
public class Person
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
}
