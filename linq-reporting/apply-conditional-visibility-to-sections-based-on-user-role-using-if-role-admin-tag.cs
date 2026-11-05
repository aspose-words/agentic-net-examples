using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider required by Aspose.Words.
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Create a new document and a builder to insert LINQ Reporting tags.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Header.
        builder.Writeln("User Report");
        builder.Writeln();

        // Show the role from the model.
        builder.Writeln("Role: <<[model.Role]>>");
        builder.Writeln();

        // Conditional section visible only for Admin role.
        builder.Writeln("<<if [model.Role == \"Admin\"]>>");
        builder.Writeln("=== Admin Section ===");
        builder.Writeln("Confidential admin data goes here.");
        builder.Writeln("<</if>>");

        // Section visible to all users.
        builder.Writeln("General information visible to all users.");

        // Save the template (optional step to demonstrate saving/loading).
        const string templatePath = "UserReportTemplate.docx";
        doc.Save(templatePath);

        // Load the template for report generation.
        Document templateDoc = new Document(templatePath);

        // Prepare the data model.
        ReportModel model = new ReportModel
        {
            Role = "Admin" // Change to "User" to hide the admin section.
        };

        // Build the report using the LINQ Reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(templateDoc, model, "model");

        // Save the generated report.
        const string outputPath = "UserReportOutput.docx";
        templateDoc.Save(outputPath);
    }
}

// Data model used by the template.
public class ReportModel
{
    public string Role { get; set; } = string.Empty;
}
