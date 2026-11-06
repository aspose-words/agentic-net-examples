using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some Aspose.Words features)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data model
        var model = new ReportModel
        {
            Email = "test@example.com"
        };

        // Create a template document programmatically
        const string templatePath = "template.docx";
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Insert template tags
        builder.Writeln("Email: <<[model.Email]>>");
        builder.Writeln("<<if [model.IsValidEmail]>>Valid Email<</if>>");
        builder.Writeln("<<if [!model.IsValidEmail]>>Invalid Email<</if>>");

        // Save the template
        doc.Save(templatePath);

        // Load the template for reporting
        var template = new Document(templatePath);

        // Configure the reporting engine
        var engine = new ReportingEngine();

        // Build the report
        engine.BuildReport(template, model, "model");

        // Save the generated report
        const string outputPath = "output.docx";
        template.Save(outputPath);
    }
}

// Public data model class
public class ReportModel
{
    // Email address to validate
    public string Email { get; set; } = string.Empty;

    // Regular expression pattern for a simple email validation
    private const string EmailPattern = @"^\S+@\S+\.\S+$";

    // Computed property that indicates whether the email matches the pattern
    public bool IsValidEmail => Regex.IsMatch(Email, EmailPattern);
}
