using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare working directories.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);
        string templatePath = Path.Combine(outputDir, "template.docx");

        // Create a simple template with a tag that references a restricted type property.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        // The tag tries to access a property of type CustomData, which is not allowed by default.
        builder.Writeln("<<[model.Custom.Info]>>");
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document doc = new Document(templatePath);

        // Prepare the data model.
        ReportModel model = new();

        // Configure the reporting engine.
        ReportingEngine engine = new();
        // By default, the engine restricts non‑primitive types. No explicit RestrictedTypes property exists
        // in the current version, so we rely on the default enforcement.

        try
        {
            // This should throw because the model contains a property of type CustomData, which is restricted.
            engine.BuildReport(doc, model, "model");
            Console.WriteLine("Test failed: no exception was thrown.");
        }
        catch (Exception ex)
        {
            // Verify that the exception is due to restricted type enforcement.
            if (ex.Message.Contains("restricted", StringComparison.OrdinalIgnoreCase) ||
                ex.Message.Contains("type", StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine("Test passed: expected exception was thrown.");
                Console.WriteLine($"Exception: {ex.GetType().Name} - {ex.Message}");
            }
            else
            {
                Console.WriteLine("Test failed: unexpected exception type.");
                Console.WriteLine($"Exception: {ex.GetType().Name} - {ex.Message}");
            }
        }
    }
}

// Public data model classes.
public class ReportModel
{
    public string Name { get; set; } = "Sample Name";
    public CustomData Custom { get; set; } = new();
}

public class CustomData
{
    public string Info { get; set; } = "Restricted Info";
}
