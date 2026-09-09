using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
}

public class Model
{
    public Person Person { get; set; } = new();
}

public class Program
{
    public static void Main()
    {
        // Create a template document. The <<restrictMembers>> tag is not required;
        // restricted types are enforced by the engine configuration.
        var templatePath = "Template.docx";
        var builder = new DocumentBuilder();
        builder.Writeln("Name: <<[model.Person.Name]>>");
        builder.Writeln("Age: <<[model.Person.Age]>>");
        builder.Document.Save(templatePath);

        // Load the template for reporting.
        var doc = new Document(templatePath);

        // Restrict access to the Person type members.
        ReportingEngine.SetRestrictedTypes(typeof(Person));

        // Prepare data.
        var model = new Model
        {
            Person = new Person { Name = "John Doe", Age = 30 }
        };

        // Configure the reporting engine.
        var engine = new ReportingEngine
        {
            Options = ReportBuildOptions.AllowMissingMembers,
            MissingMemberMessage = "Restricted"
        };

        // Build the report.
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        var outputPath = "Report.docx";
        doc.Save(outputPath);

        Console.WriteLine($"Report generated: {outputPath}");
    }
}
