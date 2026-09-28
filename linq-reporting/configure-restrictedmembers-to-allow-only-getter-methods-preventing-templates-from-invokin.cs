using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    // Only getters are exposed; setters are private to prevent template access.
    public string Name { get; private set; } = "John Doe";
    public int Age { get; private set; } = 30;
}

public class Program
{
    public static void Main()
    {
        // Sample data.
        Person person = new();

        // Create a template document with LINQ Reporting tags.
        Document template = new();
        DocumentBuilder builder = new(template);
        builder.Writeln("Person Report");
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Age: <<[person.Age]>>");

        // Save the template.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // Load the template for reporting.
        Document doc = new(templatePath);

        // Configure the ReportingEngine.
        ReportingEngine engine = new();

        // No special build options needed for this example.
        engine.Options = ReportBuildOptions.None;

        // Build the report.
        engine.BuildReport(doc, person, "person");

        // Save the generated report.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
    }
}
