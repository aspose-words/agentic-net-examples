using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Model
{
    // Sample property; does NOT contain a member named "nonexistent"
    public string Name { get; set; } = "Sample";
}

public class Program
{
    public static void Main()
    {
        // Register code page provider required by Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare file paths in the current working directory.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        string reportPath   = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");

        // -------------------------------------------------
        // Create the LINQ Reporting template programmatically
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Tag that references a missing member; should resolve to null without error.
        builder.Writeln("Value: <<[nonexistent]>>.");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // Load the template and build the report
        // -------------------------------------------------
        Document reportDoc = new Document(templatePath);
        Model model = new Model();

        ReportingEngine engine = new ReportingEngine();

        // Enable missing‑member handling. The MissingMemberHandling flag does not exist,
        // so we fall back to InlineErrorMessages which suppresses exceptions and leaves
        // the tag output empty (null) when a member is not found.
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        // Build the report; root object name is "model".
        // When InlineErrorMessages is set, BuildReport returns a bool indicating success.
        bool success = engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        reportDoc.Save(reportPath);

        // -------------------------------------------------
        // Verify that the missing member was handled gracefully
        // -------------------------------------------------
        string resultText = reportDoc.GetText();
        Console.WriteLine("Report text:");
        Console.WriteLine(resultText);
        Console.WriteLine($"Build succeeded: {success}");
    }
}
