using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    public string Name { get; set; } = "";
    public int Value { get; set; }
}

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a simple template with LINQ Reporting tags.
        string templatePath = Path.Combine(outputDir, "template.docx");
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Hello <<[model.Name]>>! Value: <<[model.Value]>>");
        templateDoc.Save(templatePath);

        // Sample data model.
        var model = new ReportModel { Name = "World", Value = 123 };

        // Build report with InlineErrorMessages enabled.
        string reportInlinePath = Path.Combine(outputDir, "report_inline.docx");
        var docInline = new Document(templatePath);
        var engineInline = new ReportingEngine
        {
            Options = ReportBuildOptions.InlineErrorMessages
        };
        bool successInline = engineInline.BuildReport(docInline, model, "model");
        docInline.Save(reportInlinePath);

        // Build report with InlineErrorMessages disabled (default options).
        string reportNoInlinePath = Path.Combine(outputDir, "report_noinline.docx");
        var docNoInline = new Document(templatePath);
        var engineNoInline = new ReportingEngine(); // Options = ReportBuildOptions.None by default.
        bool successNoInline = engineNoInline.BuildReport(docNoInline, model, "model");
        docNoInline.Save(reportNoInlinePath);

        // Compare file sizes.
        long sizeInline = new FileInfo(reportInlinePath).Length;
        long sizeNoInline = new FileInfo(reportNoInlinePath).Length;

        Console.WriteLine($"Size with InlineErrorMessages: {sizeInline} bytes");
        Console.WriteLine($"Size without InlineErrorMessages: {sizeNoInline} bytes");
        Console.WriteLine($"Difference: {sizeInline - sizeNoInline} bytes");
    }
}
