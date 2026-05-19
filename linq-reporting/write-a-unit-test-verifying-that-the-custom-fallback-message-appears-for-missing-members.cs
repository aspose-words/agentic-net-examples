using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Create a template document with a tag that references a missing member.
        string templatePath = Path.Combine(outputDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        // This tag refers to a member "Name" of an object "missing" that does not exist in the data source.
        builder.Writeln("<<[missing.Name]>>");
        templateDoc.Save(templatePath);

        // Load the template back.
        Document loadedTemplate = new Document(templatePath);

        // Configure the reporting engine to allow missing members and provide a custom fallback message.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.AllowMissingMembers;
        engine.MissingMemberMessage = "Custom fallback";

        // Build the report using an empty data source. The root name is empty because the template does not reference the root object.
        bool success = engine.BuildReport(loadedTemplate, new object(), "");

        // Verify that the fallback message appears in the generated document.
        string resultText = loadedTemplate.GetText();
        bool containsFallback = resultText.Contains(engine.MissingMemberMessage);

        Console.WriteLine(containsFallback
            ? "Test passed: custom fallback message is present."
            : "Test failed: custom fallback message is missing.");

        // Save the generated document for manual inspection.
        string resultPath = Path.Combine(outputDir, "result.docx");
        loadedTemplate.Save(resultPath);
    }
}
