using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Role of the current user.
    public string Role { get; set; } = string.Empty;
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare folders.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Create the template document.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Conditional section visible only for Admin role.
        builder.Writeln("<<if [Role == \"Admin\"]>>");
        builder.Writeln("Welcome, Administrator!");
        builder.Writeln("<</if>>");

        // Conditional section visible for non‑Admin roles.
        builder.Writeln("<<if [Role != \"Admin\"]>>");
        builder.Writeln("Welcome, User.");
        builder.Writeln("<</if>>");

        // Save the template to disk.
        string templatePath = Path.Combine(outputDir, "template.docx");
        templateDoc.Save(templatePath);

        // Load the template back (required before building the report).
        Document doc = new Document(templatePath);

        // Prepare the data model.
        ReportModel model = new ReportModel
        {
            Role = "Admin" // Change to "User" to see the other section.
        };

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string resultPath = Path.Combine(outputDir, "result.docx");
        doc.Save(resultPath);
    }
}
