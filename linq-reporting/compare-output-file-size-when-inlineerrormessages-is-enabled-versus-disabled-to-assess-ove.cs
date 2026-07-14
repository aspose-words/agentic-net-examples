using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class InlineErrorMessageSizeComparison
{
    // Simple data model used as the root object for the report.
    public class Model
    {
        // Existing property – will be displayed correctly.
        public string Name { get; set; } = "Sample Name";
        // No property named Missing – referencing this will cause a template error.
    }

    public static void Main()
    {
        // Paths for temporary files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string templatePath = Path.Combine(outputDir, "template.docx");
        string reportNoInlinePath = Path.Combine(outputDir, "report_no_inline.docx");
        string reportInlinePath = Path.Combine(outputDir, "report_inline.docx");

        // -----------------------------------------------------------------
        // 1. Create a template document programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Write a valid tag.
        builder.Writeln("Name: <<[model.Name]>>");
        // Write a tag that references a missing member – this will generate an error.
        builder.Writeln("Missing: <<[model.Missing]>>");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Build report without InlineErrorMessages (errors will throw).
        // -----------------------------------------------------------------
        Document docNoInline = new Document(templatePath);
        ReportingEngine engineNoInline = new ReportingEngine();

        try
        {
            // This will throw because the template contains an invalid reference.
            engineNoInline.BuildReport(docNoInline, new Model(), "model");
        }
        catch (Exception ex)
        {
            // Swallow the exception – the document remains unchanged (no inline errors).
            Console.WriteLine("BuildReport without InlineErrorMessages threw an exception (expected).");
            Console.WriteLine($"Exception message: {ex.Message}");
        }

        // Save the resulting document (still the original template content).
        docNoInline.Save(reportNoInlinePath);

        // -----------------------------------------------------------------
        // 3. Build report with InlineErrorMessages enabled.
        // -----------------------------------------------------------------
        Document docInline = new Document(templatePath);
        ReportingEngine engineInline = new ReportingEngine
        {
            Options = ReportBuildOptions.InlineErrorMessages
        };

        // BuildReport returns false because there were errors, but the document now contains inline messages.
        bool success = engineInline.BuildReport(docInline, new Model(), "model");
        Console.WriteLine($"BuildReport with InlineErrorMessages succeeded flag: {success}");

        // Save the document that now contains the inline error messages.
        docInline.Save(reportInlinePath);

        // -----------------------------------------------------------------
        // 4. Compare file sizes.
        // -----------------------------------------------------------------
        long sizeNoInline = new FileInfo(reportNoInlinePath).Length;
        long sizeInline = new FileInfo(reportInlinePath).Length;
        long overhead = sizeInline - sizeNoInline;

        Console.WriteLine();
        Console.WriteLine("File size comparison:");
        Console.WriteLine($"Report without InlineErrorMessages: {sizeNoInline} bytes");
        Console.WriteLine($"Report with InlineErrorMessages   : {sizeInline} bytes");
        Console.WriteLine($"Overhead introduced by InlineErrorMessages: {overhead} bytes");
    }
}
