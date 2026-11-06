using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Create a template document with an <<include>> tag that points to a non‑existent file.
        string templatePath = Path.Combine(outputDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Report start");
        // The include file does not exist; we will configure the engine to treat it as optional.
        builder.Writeln("<<include file=\"missing_file.txt\">>");
        builder.Writeln("Report end");
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document doc = new Document(templatePath);

        // Configure the reporting engine to embed inline error messages (missing includes will be ignored).
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.InlineErrorMessages;

        // Build the report. No data model is required for this example.
        bool success = engine.BuildReport(doc, new object(), "model");

        // Save the generated report.
        string resultPath = Path.Combine(outputDir, "result.docx");
        doc.Save(resultPath);

        // Indicate completion (no interactive input).
        Console.WriteLine($"Report generation {(success ? "succeeded" : "failed")}. Output saved to: {resultPath}");
    }
}
