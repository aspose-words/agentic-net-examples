using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

#nullable enable

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Sample data.
        ReportModel model = new()
        {
            Persons = new()
            {
                new Person { Name = "Alice", Age = 30, Email = "alice@example.com" },
                new Person { Name = "Bob", Age = 25, Email = null },          // Email missing
                new Person { Name = null, Age = 40, Email = "charlie@example.com" } // Name missing
            }
        };

        // Create a template document with LINQ Reporting tags.
        string templatePath = "template.docx";
        Document templateDoc = new();
        DocumentBuilder builder = new(templateDoc);

        builder.Writeln("Person Report");
        builder.Writeln("==============");
        builder.Writeln("<<foreach [p in Persons]>>");
        builder.Writeln("Name: <<[p.Name]>>");
        builder.Writeln("Age: <<[p.Age]>>");
        builder.Writeln("Email: <<[p.Email]>>");
        builder.Writeln("<</foreach>>");

        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document doc = new(templatePath);

        // Configure the reporting engine.
        ReportingEngine engine = new();
        // In newer versions the option to ignore missing members is the default.
        // If needed, you can explicitly set no special options:
        engine.Options = ReportBuildOptions.None;

        // Build the report.
        bool success = engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string outputPath = "output.docx";
        doc.Save(outputPath);

        Console.WriteLine($"Report generation {(success ? "succeeded" : "failed")}. Output saved to '{outputPath}'.");
    }
}

// Data model classes.
public class ReportModel
{
    public List<Person> Persons { get; set; } = new();
}

public class Person
{
    public string? Name { get; set; }
    public int Age { get; set; }
    public string? Email { get; set; }
}
