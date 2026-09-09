using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Person
{
    public string Name { get; set; } = "John Doe";
    // Intentionally omitted property 'Age' to trigger an inline error.
}

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        const string templatePath = "Template.docx";
        const string reportPath = "Report.docx";

        // -------------------------------------------------
        // 1. Create a template document programmatically.
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Valid tag – will be replaced with the person's name.
        builder.Writeln("Name: <<[person.Name]>>");

        // Invalid tag – property 'Age' does not exist on Person.
        // With InlineErrorMessages enabled, the engine will insert <<error>> here.
        builder.Writeln("Age: <<[person.Age]>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // 2. Load the template for reporting.
        // -------------------------------------------------
        Document loadedTemplate = new Document(templatePath);

        // -------------------------------------------------
        // 3. Configure the ReportingEngine to inline error messages.
        // -------------------------------------------------
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.InlineErrorMessages
        };

        // Sample data source.
        Person person = new Person();

        // Build the report. The method returns true if parsing succeeded.
        bool success = engine.BuildReport(loadedTemplate, person, "person");

        // -------------------------------------------------
        // 4. Save the generated report.
        // -------------------------------------------------
        loadedTemplate.Save(reportPath);

        // Output the result to the console.
        Console.WriteLine($"Report generation success flag: {success}");
        Console.WriteLine($"Report saved to: {reportPath}");
    }
}
