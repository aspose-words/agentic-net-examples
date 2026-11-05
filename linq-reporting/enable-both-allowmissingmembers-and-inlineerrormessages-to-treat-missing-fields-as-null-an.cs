using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some encodings)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for template and output
        string templatePath = "template.docx";
        string outputPath = "output.docx";

        // Create template document with LINQ Reporting tags
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Write a heading
        builder.Writeln("LINQ Reporting Example with AllowMissingMembers and InlineErrorMessages");
        builder.Writeln();

        // Correct tag
        builder.Writeln("Name: <<[person.Name]>>");
        builder.Writeln("Age: <<[person.Age]>>");

        // Missing member tag (property does not exist in the model)
        builder.Writeln("Missing field (should be treated as null): <<[person.MissingField]>>");

        // Syntax error tag (invalid switch)
        builder.Writeln("Syntax error example: <<[person.Name] -invalidSwitch>>");

        // Save the template
        templateDoc.Save(templatePath);

        // Load the template for reporting
        Document doc = new Document(templatePath);

        // Prepare data model
        Person person = new Person
        {
            Name = "John Doe",
            Age = 30
        };

        // Configure reporting engine with required options
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.AllowMissingMembers | ReportBuildOptions.InlineErrorMessages;

        // Build the report
        bool success = engine.BuildReport(doc, person, "person");

        // Save the generated report
        doc.Save(outputPath);

        // Output result status
        Console.WriteLine($"Report generation success: {success}");
        Console.WriteLine($"Output saved to: {Path.GetFullPath(outputPath)}");
    }
}

// Data model class
public class Person
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
}
