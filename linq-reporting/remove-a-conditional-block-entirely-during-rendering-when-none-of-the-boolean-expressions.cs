using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    public bool ShowFirst { get; set; } = false;
    public bool ShowSecond { get; set; } = false;
}

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create the template document.
        string templatePath = Path.Combine(outputDir, "template.docx");
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Static content before the conditional block.
        builder.Writeln("Report Header");
        builder.Writeln();

        // Conditional block that should be removed when both conditions are false.
        builder.Writeln("<<if [model.ShowFirst || model.ShowSecond]>>");
        builder.Writeln("=== Conditional Section Start ===");
        builder.Writeln("<<if [model.ShowFirst]>>First condition is true.<</if>>");
        builder.Writeln("<<if [model.ShowSecond]>>Second condition is true.<</if>>");
        builder.Writeln("=== Conditional Section End ===");
        builder.Writeln("<</if>>");

        // Static content after the conditional block.
        builder.Writeln();
        builder.Writeln("Report Footer");

        // Save the template.
        templateDoc.Save(templatePath);

        // Load the template for reporting.
        Document doc = new Document(templatePath);

        // Prepare the data model with both conditions false.
        ReportModel model = new ReportModel
        {
            ShowFirst = false,
            ShowSecond = false
        };

        // Build the report.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Save the generated report.
        string outputPath = Path.Combine(outputDir, "output.docx");
        doc.Save(outputPath);
    }
}
