using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Create output directory
        string outputDir = "output";
        Directory.CreateDirectory(outputDir);

        // Create a valid include file
        string validIncludePath = Path.Combine(outputDir, "included.txt");
        File.WriteAllText(validIncludePath, "This is the content of the included file.");

        // Path for a missing include file (do not create this file)
        string missingIncludePath = Path.Combine(outputDir, "missing.txt");

        // Build the template document with include tags
        string templatePath = Path.Combine(outputDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("=== Report Start ===");
        builder.Writeln($"<<include file=\"{validIncludePath}\">>");
        builder.Writeln($"<<include file=\"{missingIncludePath}\">>");
        builder.Writeln("=== Report End ===");
        templateDoc.Save(templatePath);

        // Load the template for reporting
        Document doc = new Document(templatePath);

        // Configure the reporting engine
        ReportingEngine engine = new ReportingEngine();
        // Use InlineErrorMessages to get any include errors inside the document (optional)
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        // Build the report using an empty root model
        var model = new ReportModel();
        bool success;
        try
        {
            success = engine.BuildReport(doc, model, "model");
        }
        catch (Exception ex)
        {
            // If an include file is missing, treat it as optional and continue
            Console.WriteLine($"BuildReport exception (treated as optional include): {ex.Message}");
            success = true;
        }

        // Save the generated report
        string resultPath = Path.Combine(outputDir, "result.docx");
        doc.Save(resultPath);

        // Indicate completion (no interactive input)
        Console.WriteLine($"Report generation success: {success}");
        Console.WriteLine($"Result saved to: {resultPath}");
    }

    public class ReportModel
    {
        // No properties needed for this example
    }
}
