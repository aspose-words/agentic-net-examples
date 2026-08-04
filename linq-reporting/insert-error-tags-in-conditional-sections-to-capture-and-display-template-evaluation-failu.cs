using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;
using Newtonsoft.Json;
using System.Text;

public class Program
{
    public static void Main()
    {
        // Prepare file paths in the current working directory.
        string workDir = Directory.GetCurrentDirectory();
        string templatePath = Path.Combine(workDir, "template.docx");
        string outputPath = Path.Combine(workDir, "report.docx");

        // 1. Create a blank template document and insert LINQ Reporting tags.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("=== Report Start ===");
        // Conditional block that references a non‑existent member.
        builder.Writeln("<<if [model.ShowDetails]>>");
        builder.Writeln("Attempting to read missing member: <<[model.MissingProperty]>>");
        builder.Writeln("<</if>>");
        builder.Writeln("=== Report End ===");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // 2. Load the saved template (simulating an external file).
        Document doc = new Document(templatePath);

        // 3. Prepare the data model.
        ReportModel model = new ReportModel
        {
            ShowDetails = true
            // Detail is intentionally left empty; the template references a property that does not exist.
        };

        // 4. Configure the ReportingEngine to inline error messages.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        // 5. Build the report. Missing members will be replaced by inline error messages.
        bool success = engine.BuildReport(doc, model, "model");

        // 6. Save the generated report.
        doc.Save(outputPath);

        // 7. Output results to the console.
        Console.WriteLine($"Report generation success flag: {success}");
        Console.WriteLine("Generated document text:");
        Console.WriteLine(doc.GetText());
    }
}

// Public data model aligned with the template.
public class ReportModel
{
    // The conditional checks this flag.
    public bool ShowDetails { get; set; } = false;

    // This property exists but is not used in the faulty tag.
    public string Detail { get; set; } = "";
}
