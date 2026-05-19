using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Model
{
    // Sample property; the template will reference a non‑existent member.
    public string Name { get; set; } = "Sample";
}

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.docx");
        string reportPath   = Path.Combine(Directory.GetCurrentDirectory(), "Report.docx");

        // ---------- Create the template ----------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a LINQ Reporting tag that references a missing member.
        builder.Writeln("<<[model.Nonexistent]>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // ---------- Load the template ----------
        Document loadedTemplate = new Document(templatePath);

        // ---------- Prepare data ----------
        Model model = new Model(); // No 'Nonexistent' property.

        // ---------- Configure the reporting engine ----------
        ReportingEngine engine = new ReportingEngine
        {
            Options = ReportBuildOptions.AllowMissingMembers
            // MissingMemberMessage left as default (empty) so missing members become empty strings.
        };

        // Build the report. The missing member should be treated as null (empty) without throwing.
        bool success = engine.BuildReport(loadedTemplate, model, "model");

        // Save the generated report.
        loadedTemplate.Save(reportPath);

        // ---------- Verify the result ----------
        string resultText = loadedTemplate.GetText();

        // If the missing member was handled correctly, the placeholder is replaced with an empty string.
        bool isHandled = string.IsNullOrWhiteSpace(resultText);

        Console.WriteLine(isHandled
            ? "Missing member handled as null."
            : "Unexpected content found in the report.");

        // Optionally, indicate whether the build succeeded.
        Console.WriteLine($"BuildReport success flag: {success}");
    }
}
