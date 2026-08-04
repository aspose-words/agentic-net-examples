using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    // Simple data model with only public properties.
    public class Person
    {
        public string Name { get; set; } = "John Doe";
        public int Age { get; set; } = 30;
        // Private field and method are not exposed to the template.
        private string Secret = "Hidden";
        private string GetSecret() => Secret;
    }

    public static void Main()
    {
        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // 1. Create the template document programmatically.
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("Person Report");
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Age: <<[person.Age]>>");
        // Attempting to access a non‑public member would fail, but we don't include such a tag.

        // Save the template to disk.
        string templatePath = Path.Combine(outputDir, "Template.docx");
        template.Save(templatePath);

        // 2. Load the template back (required before building the report).
        Document loadedTemplate = new Document(templatePath);

        // 3. Restrict access to types that should not be used in templates.
        // Here we restrict the System.Environment type as an example.
        // This does not affect our Person type, which only exposes public properties.
        ReportingEngine.SetRestrictedTypes(typeof(Environment));

        // 4. Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        // No special options are needed for this simple example.
        Person person = new Person();
        engine.BuildReport(loadedTemplate, person, "person");

        // 5. Save the generated report.
        string reportPath = Path.Combine(outputDir, "Report.docx");
        loadedTemplate.Save(reportPath);

        // Indicate completion (no interactive prompts).
        Console.WriteLine($"Report generated at: {reportPath}");
    }
}
