using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data model.
        var model = new Person
        {
            Name = "John Doe",
            Email = "john.doe@example.com"
        };

        // Create a template document programmatically.
        var templatePath = "Template.docx";
        CreateTemplate(templatePath);

        // Load the template.
        var doc = new Document(templatePath);

        // Configure the reporting engine.
        var engine = new ReportingEngine();
        // Register System.Text.RegularExpressions.Regex to allow its static members in template expressions.
        engine.KnownTypes.Add(typeof(Regex));

        // Build the report using the model as the root object named "model".
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        var outputPath = "Output.docx";
        doc.Save(outputPath);
    }

    private static void CreateTemplate(string filePath)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Insert placeholders that will be replaced by the reporting engine.
        builder.Writeln("Name: <<[model.Name]>>");
        builder.Writeln("Email: <<[model.Email]>>");
        // Use Regex.IsMatch to validate the email format directly in the template.
        builder.Writeln("Is Email Valid? <<[Regex.IsMatch(model.Email, \"^\\\\S+@\\\\S+\\\\.\\\\S+$\")]>>");

        // Save the template to disk.
        doc.Save(filePath);
    }
}

// Simple data model with public properties.
public class Person
{
    public string Name { get; set; } = string.Empty;
    public string Email { get; set; } = string.Empty;
}
