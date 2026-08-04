using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class ReportModel
{
    // Initialize to avoid nullable warnings.
    public string Name { get; set; } = string.Empty;
}

public class Program
{
    public static void Main()
    {
        // Paths for the template and generated reports.
        string templatePath = Path.Combine(Environment.CurrentDirectory, "Template.docx");
        string reportWithInlinePath = Path.Combine(Environment.CurrentDirectory, "Report_WithInline.docx");
        string reportWithoutInlinePath = Path.Combine(Environment.CurrentDirectory, "Report_WithoutInline.docx");

        // -----------------------------------------------------------------
        // 1. Create a simple template document with a LINQ Reporting tag.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);
        builder.Writeln("Hello <<[model.Name]>>!"); // LINQ Reporting tag.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Prepare sample data.
        // -----------------------------------------------------------------
        ReportModel model = new ReportModel { Name = "World" };

        // -----------------------------------------------------------------
        // 3. Generate report with InlineErrorMessages enabled.
        // -----------------------------------------------------------------
        Document docWithInline = new Document(templatePath);
        ReportingEngine engineWithInline = new ReportingEngine
        {
            Options = ReportBuildOptions.InlineErrorMessages
        };
        bool successWithInline = engineWithInline.BuildReport(docWithInline, model, "model");
        // Save the generated document.
        docWithInline.Save(reportWithInlinePath);

        // -----------------------------------------------------------------
        // 4. Generate report with InlineErrorMessages disabled (default options).
        // -----------------------------------------------------------------
        Document docWithoutInline = new Document(templatePath);
        ReportingEngine engineWithoutInline = new ReportingEngine(); // No InlineErrorMessages flag.
        bool successWithoutInline = engineWithoutInline.BuildReport(docWithoutInline, model, "model");
        docWithoutInline.Save(reportWithoutInlinePath);

        // -----------------------------------------------------------------
        // 5. Compare file sizes and output the results.
        // -----------------------------------------------------------------
        long sizeWithInline = new FileInfo(reportWithInlinePath).Length;
        long sizeWithoutInline = new FileInfo(reportWithoutInlinePath).Length;

        Console.WriteLine($"Report with InlineErrorMessages:   Size = {sizeWithInline} bytes, Build success = {successWithInline}");
        Console.WriteLine($"Report without InlineErrorMessages: Size = {sizeWithoutInline} bytes, Build success = {successWithoutInline}");
        Console.WriteLine($"Size overhead introduced by InlineErrorMessages: {sizeWithInline - sizeWithoutInline} bytes");
    }
}
