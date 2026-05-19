using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Model
{
    public string Name { get; set; } = "John Doe";
}

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some environments).
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        string templatePath = Path.Combine(outputDir, "template.docx");
        string resultPath = Path.Combine(outputDir, "result.docx");

        // Create a template document with a correct tag and an intentional error tag.
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Name: <<[model.Name]>>");
        builder.Writeln("Missing: <<[model.Missing]>>"); // This member does not exist.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document doc = new Document(templatePath);

        // Prepare the data source.
        Model model = new Model();

        // Configure the reporting engine to inline error messages.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        // Build the report; the returned flag indicates if parsing succeeded.
        bool success = engine.BuildReport(doc, model, "model");

        // Save the generated report.
        doc.Save(resultPath);

        // Output the result status.
        Console.WriteLine($"Report build success: {success}");
        Console.WriteLine($"Report saved to: {resultPath}");
    }
}
