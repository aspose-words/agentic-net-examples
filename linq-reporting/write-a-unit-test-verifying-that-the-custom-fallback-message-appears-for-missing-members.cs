using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Model
{
    // Existing property used in the template.
    public string Existing { get; set; } = "World";
    // No property named Missing – this will trigger the fallback message.
}

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // Create a template document that references an existing and a missing member.
        // -----------------------------------------------------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);
        builder.Writeln("Hello <<[model.Existing]>> and <<[model.Missing]>>!");

        // -----------------------------------------------------------------
        // Prepare the data model.
        // -----------------------------------------------------------------
        Model model = new Model();

        // -----------------------------------------------------------------
        // Configure the reporting engine to embed inline error messages.
        // -----------------------------------------------------------------
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        // -----------------------------------------------------------------
        // Build the report. The method returns false because a missing member is encountered.
        // -----------------------------------------------------------------
        bool success = engine.BuildReport(template, model, "model");

        // -----------------------------------------------------------------
        // Verify that an inline error message appears in the generated document.
        // The default inline error message contains the word "Error".
        // -----------------------------------------------------------------
        string resultText = template.GetText();
        bool containsErrorMessage = resultText.Contains("Error");
        bool testPassed = !success && containsErrorMessage;

        Console.WriteLine(testPassed
            ? "Test passed: inline error message correctly inserted."
            : "Test failed: inline error message not found or build succeeded unexpectedly.");

        // -----------------------------------------------------------------
        // Save the generated document for manual inspection.
        // -----------------------------------------------------------------
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "Report.docx");
        template.Save(outputPath);
    }
}
