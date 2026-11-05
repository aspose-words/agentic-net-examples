using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare working directories.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "LinqReportingExample");
        Directory.CreateDirectory(workDir);
        string templatePath = Path.Combine(workDir, "template.docx");
        string outputPath = Path.Combine(workDir, "output.docx");

        // -----------------------------------------------------------------
        // 1. Create the template document with a missing member tag.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Existing: <<[model.Existing]>>; Missing: <<[model.Missing]>>.");
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Define the data model (Missing property is intentionally absent).
        // -----------------------------------------------------------------
        var model = new ReportModel
        {
            Existing = "Value"
            // No Missing property.
        };

        // -----------------------------------------------------------------
        // 3. Load the template and build the report with AllowMissingMembers.
        // -----------------------------------------------------------------
        Document doc = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.AllowMissingMembers;
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save(outputPath);

        // -----------------------------------------------------------------
        // 4. Verify that the missing member was rendered as an empty string.
        // -----------------------------------------------------------------
        string resultText = doc.GetText();
        Console.WriteLine("Generated document text:");
        Console.WriteLine(resultText.Trim());

        bool testPassed = resultText.Contains("Existing: Value") && resultText.Contains("Missing: .");
        Console.WriteLine(testPassed ? "Test passed: missing member rendered as null literal." : "Test failed.");
    }

    // Public data model class.
    public class ReportModel
    {
        // Existing member used in the template.
        public string Existing { get; set; } = string.Empty;
        // No Missing property – this is intentional for the test.
    }
}
