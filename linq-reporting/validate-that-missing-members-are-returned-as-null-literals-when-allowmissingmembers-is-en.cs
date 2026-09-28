using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Ensure the output directory exists.
        string outputDir = "output";
        Directory.CreateDirectory(outputDir);

        // Create a template document with LINQ Reporting tags.
        string templatePath = Path.Combine(outputDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Write a line that references an existing member.
        builder.Writeln("Existing: <<[model.Existing]>>");
        // Write a line that references a missing member.
        builder.Writeln("Missing: <<[model.Missing]>>");
        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document doc = new Document(templatePath);

        // Prepare the data model with only the Existing property.
        var model = new ReportModel
        {
            Existing = "Hello World"
            // Note: Missing property is intentionally not defined.
        };

        // Configure the reporting engine to allow missing members.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.AllowMissingMembers;

        // Build the report.
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string resultPath = Path.Combine(outputDir, "result.docx");
        doc.Save(resultPath);

        // Output the resulting document text to the console.
        Console.WriteLine("Generated document text:");
        Console.WriteLine(doc.GetText());
    }

    // Public data model class.
    public class ReportModel
    {
        // Existing member used in the template.
        public string Existing { get; set; } = string.Empty;
        // No Missing member defined; it will be treated as null when AllowMissingMembers is enabled.
    }
}
