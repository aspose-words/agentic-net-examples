using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    public string Role { get; set; } = string.Empty;
    public string UserName { get; set; } = string.Empty;
}

public class Program
{
    public static void Main()
    {
        // Paths for the template and generated reports.
        const string templatePath = "template.docx";
        const string adminReportPath = "report_admin.docx";
        const string userReportPath = "report_user.docx";

        // -------------------------------------------------
        // 1. Create the template document programmatically.
        // -------------------------------------------------
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);

        // Always visible content.
        builder.Writeln("User: <<[model.UserName]>>");
        builder.Writeln("Role: <<[model.Role]>>");

        // Conditional section visible only for Admin role.
        builder.Writeln("<<if [model.Role == \"Admin\"]>>");
        builder.Writeln("=== Admin Section ===");
        builder.Writeln("Confidential information displayed only to administrators.");
        builder.Writeln("<</if>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // 2. Load the template and generate a report for an Admin.
        // -------------------------------------------------
        var adminDoc = new Document(templatePath);
        var adminModel = new ReportModel
        {
            UserName = "Alice",
            Role = "Admin"
        };

        var engine = new ReportingEngine
        {
            Options = ReportBuildOptions.RemoveEmptyParagraphs
        };
        engine.BuildReport(adminDoc, adminModel, "model");
        adminDoc.Save(adminReportPath);

        // -------------------------------------------------
        // 3. Load the template and generate a report for a non‑Admin user.
        // -------------------------------------------------
        var userDoc = new Document(templatePath);
        var userModel = new ReportModel
        {
            UserName = "Bob",
            Role = "User"
        };

        engine.BuildReport(userDoc, userModel, "model");
        userDoc.Save(userReportPath);
    }
}
