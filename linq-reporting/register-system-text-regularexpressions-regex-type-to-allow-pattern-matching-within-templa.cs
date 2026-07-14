using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Sample data model.
        var person = new Person
        {
            Email = "john.doe@example.com"
        };

        // Create a blank document and a builder to insert template tags.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Simple field showing the email address.
        builder.Writeln("Email address: <<[person.Email]>>");

        // Validation tag using Regex.IsMatch.
        // The pattern must contain a double backslash so that the engine receives a literal "\.".
        builder.Writeln(
            "Validation result: <<[Regex.IsMatch(person.Email, \"^[^@]+@[^@]+\\\\.[^@]+$\") ? \"Valid\" : \"Invalid\"]>>");

        // Configure the LINQ Reporting engine.
        var engine = new ReportingEngine();

        // Register the Regex type to allow static member access in template expressions.
        engine.KnownTypes.Add(typeof(Regex));

        // Build the report using the template and the data source.
        engine.BuildReport(doc, person, "person");

        // Save the generated document.
        const string outputPath = "Report.docx";
        doc.Save(outputPath);
        Console.WriteLine($"Report generated: {Path.GetFullPath(outputPath)}");
    }
}

// Simple data model with a non‑nullable Email property.
public class Person
{
    public string Email { get; set; } = string.Empty;
}
