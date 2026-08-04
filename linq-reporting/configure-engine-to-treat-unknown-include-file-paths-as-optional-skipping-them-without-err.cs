using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Create a temporary folder for the example files.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "ExampleFiles");
        Directory.CreateDirectory(workDir);

        // Paths for the template and the generated report.
        string templatePath = Path.Combine(workDir, "template.docx");
        string resultPath = Path.Combine(workDir, "result.docx");

        // -----------------------------------------------------------------
        // 1. Build a simple template that contains a <<doc>> tag pointing
        //    to a non‑existent file. The engine will be configured to treat this
        //    include as optional and skip it without throwing an error.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("=== Report Start ===");
        // The file "MissingFile.docx" does not exist.
        // Use the supported <<doc>> tag; when the file is missing the engine will
        // output nothing (treated as optional) because we enable the
        // AllowMissingMembers flag and remove empty paragraphs.
        builder.Writeln("<<doc [MissingFile.docx]>>");
        builder.Writeln("=== Report End ===");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and configure the ReportingEngine.
        // -----------------------------------------------------------------
        Document loadedTemplate = new Document(templatePath);

        ReportingEngine engine = new ReportingEngine();

        // Allow missing members (not used here but required for the option) and
        // remove empty paragraphs that may be left after the missing include.
        engine.Options = ReportBuildOptions.AllowMissingMembers | ReportBuildOptions.RemoveEmptyParagraphs;
        engine.MissingMemberMessage = string.Empty; // No placeholder text for missing members.

        // Build the report. No data source is required for this example.
        engine.BuildReport(loadedTemplate, new object());

        // Save the generated report.
        loadedTemplate.Save(resultPath);

        // Inform the user where the files are located (no interactive prompts).
        Console.WriteLine($"Template saved to: {templatePath}");
        Console.WriteLine($"Report generated at: {resultPath}");
    }
}
