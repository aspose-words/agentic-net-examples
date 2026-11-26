using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Register code pages (required for some Aspose.Words features)
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Prepare sample data model
        Person person = new()
        {
            Name = "John Doe",
            Age = 30,
            Secret = "TopSecret"
        };

        // Create a template document programmatically
        const string templatePath = "Template.docx";
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);

        builder.Writeln("Person Report");
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Age: <<[person.Age]>>");
        // The secret field is intentionally omitted to demonstrate restricted access.
        // builder.Writeln("Secret (should be restricted): <<[person.Secret]>>");

        templateDoc.Save(templatePath);

        // Load the template for reporting
        Document reportDoc = new(templatePath);

        // Configure the ReportingEngine (no special restrictions needed; only public members are accessible)
        ReportingEngine engine = new();

        // Build the report
        bool success = engine.BuildReport(reportDoc, person, "person");

        // Save the generated report
        const string outputPath = "Report.docx";
        reportDoc.Save(outputPath);

        // Output result status
        Console.WriteLine($"Report generation {(success ? "succeeded" : "failed")}. Output saved to '{outputPath}'.");
    }
}

// Public data model with public properties
public class Person
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
    public string Secret { get; set; } = string.Empty;
}
