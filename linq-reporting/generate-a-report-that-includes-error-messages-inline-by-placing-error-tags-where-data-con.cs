using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some Aspose.Words features)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare output directory
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a template document programmatically
        string templatePath = Path.Combine(outputDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("=== Customer Report ===");
        builder.Writeln("Name: <<[model.Name]>>");
        builder.Writeln("Age: <<[model.Age]>>");
        // This tag references a non‑existent property and will trigger an inline error message
        builder.Writeln("Missing Property: <<[model.NonExistent]>>");
        builder.Writeln("=== End of Report ===");

        templateDoc.Save(templatePath);

        // Load the template (simulating a real scenario where the template is read from disk)
        Document doc = new Document(templatePath);

        // Prepare the data model
        ReportModel model = new ReportModel
        {
            Name = "John Doe",
            Age = 30
        };

        // Configure the reporting engine to embed inline error messages
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        // Build the report
        bool success = engine.BuildReport(doc, model, "model");

        // Save the generated report
        string resultPath = Path.Combine(outputDir, "output.docx");
        doc.Save(resultPath);

        // Output the result status
        Console.WriteLine($"Report generation success: {success}");
        Console.WriteLine($"Template saved to: {templatePath}");
        Console.WriteLine($"Report saved to: {resultPath}");
    }
}

// Data model used by the report
public class ReportModel
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
}
