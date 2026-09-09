using System;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare sample data.
        var model = new Person
        {
            Email = "test@example.com"
        };

        // -----------------------------------------------------------------
        // 1. Create a template document programmatically.
        // -----------------------------------------------------------------
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Insert a line that shows the raw email address.
        builder.Writeln("Email: <<[model.Email]>>");

        // Insert a line that validates the email using Regex.IsMatch.
        // The regular expression is escaped for C# string literals.
        builder.Writeln(
            "IsValid: <<[Regex.IsMatch(model.Email, \"^\\\\S+@\\\\S+\\\\.\\\\S+$\")]>>");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template document (required before building the report).
        // -----------------------------------------------------------------
        var loadedTemplate = new Document(templatePath);

        // -----------------------------------------------------------------
        // 3. Configure the ReportingEngine.
        // -----------------------------------------------------------------
        var engine = new ReportingEngine();

        // Register the Regex type so that static members can be used in template expressions.
        engine.KnownTypes.Add(typeof(Regex));

        // -----------------------------------------------------------------
        // 4. Build the report.
        // -----------------------------------------------------------------
        // The root object name used in the template is "model".
        engine.BuildReport(loadedTemplate, model, "model");

        // -----------------------------------------------------------------
        // 5. Save the generated report.
        // -----------------------------------------------------------------
        const string reportPath = "Report.docx";
        loadedTemplate.Save(reportPath);
    }
}

// Simple data model with a non‑nullable Email property.
public class Person
{
    public string Email { get; set; } = string.Empty;
}
