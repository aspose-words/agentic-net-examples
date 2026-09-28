using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Role of the current user.
    public string Role { get; set; } = string.Empty;
    // Additional content (not used in this example but shown for completeness).
    public string Content { get; set; } = string.Empty;
}

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        string templatePath = "Template.docx";
        string outputPath = "Report_Output.docx";

        // -----------------------------------------------------------------
        // Step 1: Create the template document with conditional tags.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("=== Report Header ===");
        // Conditional section: visible only when Role == "Admin".
        builder.Writeln("<<if [model.Role == \"Admin\"]>>");
        builder.Writeln("Admin-only section: Confidential information displayed here.");
        builder.Writeln("<</if>>");
        builder.Writeln("=== Report Footer ===");

        // Save the template to disk before building the report.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // Step 2: Prepare the data model.
        // -----------------------------------------------------------------
        ReportModel model = new ReportModel
        {
            Role = "Admin", // Change to other values (e.g., "User") to hide the section.
            Content = "Sample content"
        };

        // -----------------------------------------------------------------
        // Step 3: Load the template and generate the report.
        // -----------------------------------------------------------------
        Document docToReport = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();

        // Build the report using the model as the root object named "model".
        engine.BuildReport(docToReport, model, "model");

        // Save the generated report.
        docToReport.Save(outputPath);
    }
}
