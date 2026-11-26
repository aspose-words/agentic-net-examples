using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words if needed.
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare folders.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a template document with a conditional signature block.
        string templatePath = Path.Combine(outputDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        builder.Writeln("=== Sample Report ===");
        builder.Writeln("");
        // Conditional block: include signature only when Signatory is not null.
        builder.Writeln("<<if [Signatory != null]>>");
        builder.Writeln("Signature: <<[Signatory]>>");
        builder.Writeln("<</if>>");

        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document reportDoc = new Document(templatePath);

        // Prepare the data model.
        ReportModel model = new()
        {
            Signatory = "John Doe"
        };

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;
        engine.BuildReport(reportDoc, model, "model");

        // Save the generated report.
        string resultPath = Path.Combine(outputDir, "output.docx");
        reportDoc.Save(resultPath);
    }
}

// Data model for the report.
public class ReportModel
{
    // Nullable property; when null, the signature block will be omitted.
    public string? Signatory { get; set; }
}
