using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for template and reports.
        string templatePath = Path.Combine(outputDir, "template.docx");
        string reportNoInlinePath = Path.Combine(outputDir, "report_no_inline.docx");
        string reportInlinePath = Path.Combine(outputDir, "report_inline.docx");

        // Create the template document with a valid tag and a missing tag.
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Customer Name: <<[model.Name]>>");
        builder.Writeln("Missing Field: <<[model.MissingProperty]>>"); // This property does not exist.
        templateDoc.Save(templatePath);

        // Sample data model (MissingProperty is intentionally omitted).
        var model = new ReportModel { Name = "John Doe" };

        // Generate report without InlineErrorMessages.
        var reportNoInline = new Document(templatePath);
        var engineNoInline = new ReportingEngine();
        engineNoInline.Options = ReportBuildOptions.None;
        bool successNoInline;
        try
        {
            successNoInline = engineNoInline.BuildReport(reportNoInline, model, "model");
        }
        catch (InvalidOperationException)
        {
            // BuildReport throws when a missing property is encountered and InlineErrorMessages is disabled.
            successNoInline = false;
        }
        reportNoInline.Save(reportNoInlinePath);

        // Generate report with InlineErrorMessages.
        var reportInline = new Document(templatePath);
        var engineInline = new ReportingEngine();
        engineInline.Options = ReportBuildOptions.InlineErrorMessages;
        bool successInline = engineInline.BuildReport(reportInline, model, "model");
        reportInline.Save(reportInlinePath);

        // Compare file sizes.
        long sizeNoInline = new FileInfo(reportNoInlinePath).Length;
        long sizeInline = new FileInfo(reportInlinePath).Length;

        Console.WriteLine($"Report without InlineErrorMessages size: {sizeNoInline} bytes (Success: {successNoInline})");
        Console.WriteLine($"Report with InlineErrorMessages size:    {sizeInline} bytes (Success: {successInline})");
        Console.WriteLine($"Size overhead: {sizeInline - sizeNoInline} bytes");
    }
}

// Data model used for the report.
public class ReportModel
{
    public string Name { get; set; } = "";
}
