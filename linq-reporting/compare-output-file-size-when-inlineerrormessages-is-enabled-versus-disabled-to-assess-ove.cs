using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for Aspose.Words).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Prepare sample data model.
        var model = new ReportModel
        {
            Name = "John Doe"
        };

        // Create a template document programmatically.
        const string templatePath = "Template.docx";
        var builder = new DocumentBuilder();
        builder.Writeln("Customer Name: <<[model.Name]>>");
        builder.Writeln("This line contains an invalid tag to trigger an error:");
        builder.Writeln("<<[model.NonExistent]>>"); // Invalid expression.
        builder.Document.Save(templatePath);

        // Load the template for reporting.
        var template = new Document(templatePath);

        // ---------- Report without InlineErrorMessages ----------
        var reportNoInline = (Document)template.Clone();
        var engineNoInline = new ReportingEngine(); // Default options.

        bool successNoInline;
        try
        {
            successNoInline = engineNoInline.BuildReport(reportNoInline, model, "model");
        }
        catch (Exception ex)
        {
            // BuildReport throws when an expression cannot be resolved and InlineErrorMessages is disabled.
            Console.WriteLine($"BuildReport (no inline) threw an exception: {ex.Message}");
            successNoInline = false;
        }

        const string outputNoInline = "Report_NoInline.docx";
        reportNoInline.Save(outputNoInline);

        // ---------- Report with InlineErrorMessages ----------
        var reportInline = (Document)template.Clone();
        var engineInline = new ReportingEngine();
        engineInline.Options = ReportBuildOptions.InlineErrorMessages; // Enable inline error messages.

        bool successInline = engineInline.BuildReport(reportInline, model, "model");
        const string outputInline = "Report_Inline.docx";
        reportInline.Save(outputInline);

        // Compare file sizes.
        long sizeNoInline = new FileInfo(outputNoInline).Length;
        long sizeInline = new FileInfo(outputInline).Length;

        Console.WriteLine($"Report without InlineErrorMessages size: {sizeNoInline} bytes");
        Console.WriteLine($"Report with InlineErrorMessages size:    {sizeInline} bytes");
        Console.WriteLine($"Size overhead: {sizeInline - sizeNoInline} bytes");
        Console.WriteLine($"Build success without inline: {successNoInline}");
        Console.WriteLine($"Build success with inline:    {successInline}");
    }
}

// Public data model class.
public class ReportModel
{
    public string Name { get; set; } = string.Empty;
}
