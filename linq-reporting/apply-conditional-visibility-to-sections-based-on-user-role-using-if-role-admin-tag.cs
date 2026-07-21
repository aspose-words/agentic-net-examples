using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Role of the user; set to "Admin" to show the conditional section.
    public string Role { get; set; } = string.Empty;

    // Additional data that could be used in the template.
    public string Message { get; set; } = string.Empty;
}

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create the template document programmatically.
        // -----------------------------------------------------------------
        var template = new Document();
        var builder = new DocumentBuilder(template);

        // Static content that is always visible.
        builder.Writeln("=== Report Header ===");
        builder.Writeln("User Role: <<[model.Role]>>");
        builder.Writeln();

        // Conditional section: visible only when Role == "Admin".
        builder.Writeln("<<if [model.Role == \"Admin\"]>>");
        builder.Writeln(">>> This section is visible only to administrators.");
        builder.Writeln(">>> Message: <<[model.Message]>>");
        builder.Writeln("<</if>>");

        // Footer content.
        builder.Writeln();
        builder.Writeln("=== Report Footer ===");

        // Save the template to disk.
        const string templatePath = "Template.docx";
        template.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report.
        // -----------------------------------------------------------------
        var loadedTemplate = new Document(templatePath);

        // Sample data model.
        var model = new ReportModel
        {
            Role = "Admin",               // Change to other values to hide the conditional block.
            Message = "Welcome to the admin dashboard."
        };

        // Configure and execute the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None; // No special options needed for this example.
        engine.BuildReport(loadedTemplate, model, "model");

        // Save the generated report.
        const string reportPath = "Report.docx";
        loadedTemplate.Save(reportPath);
    }
}
