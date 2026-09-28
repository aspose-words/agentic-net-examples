using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Initialize to avoid nullable warnings.
    public string Status { get; set; } = string.Empty;
}

public class Program
{
    public static void Main()
    {
        // Ensure output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a template document with conditional tags.
        string templatePath = Path.Combine(outputDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Conditional blocks using double‑quoted string literals (required by LINQ Reporting syntax).
        builder.Writeln("<<if [model.Status == \"Approved\"]>>Approved<</if>>");
        builder.Writeln("<<if [model.Status == \"Pending\"]>>Pending<</if>>");
        // Default block when none of the above conditions are true.
        builder.Writeln("<<if [model.Status != \"Approved\" && model.Status != \"Pending\"]>>Unknown<</if>>");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document doc = new Document(templatePath);

        // Sample data where no condition matches to trigger the default block.
        ReportModel model = new ReportModel { Status = "Other" };

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string resultPath = Path.Combine(outputDir, "result.docx");
        doc.Save(resultPath);

        Console.WriteLine($"Report generated: {resultPath}");
    }
}
