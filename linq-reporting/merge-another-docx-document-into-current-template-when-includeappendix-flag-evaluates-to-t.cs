using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider for possible encoding needs
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // Create output directory
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "output");
        Directory.CreateDirectory(outputDir);

        // Paths for template, appendix, and final report
        string templatePath = Path.Combine(outputDir, "template.docx");
        string appendixPath = Path.Combine(outputDir, "appendix.docx");
        string resultPath = Path.Combine(outputDir, "report.docx");

        // 1. Build the main template document
        var templateDoc = new Document();
        var builder = new DocumentBuilder(templateDoc);
        builder.Writeln("=== Report ===");
        builder.Writeln("Title: <<[model.Title]>>");
        builder.Writeln("Include Appendix: <<[model.IncludeAppendix]>>");
        // Bookmark where the appendix will be inserted
        builder.StartBookmark("AppendixHere");
        builder.Writeln("<<[model.AppendixPlaceholder]>>");
        builder.EndBookmark("AppendixHere");
        templateDoc.Save(templatePath);

        // 2. Build the appendix document
        var appendixDoc = new Document();
        var appendixBuilder = new DocumentBuilder(appendixDoc);
        appendixBuilder.Writeln("=== Appendix ===");
        appendixBuilder.Writeln("This is the appendix content.");
        appendixDoc.Save(appendixPath);

        // 3. Prepare the data model
        var model = new ReportModel
        {
            Title = "Quarterly Sales Report",
            IncludeAppendix = true,
            AppendixPlaceholder = "[Appendix will be inserted here]"
        };

        // 4. Load the template for reporting
        var reportDoc = new Document(templatePath);

        // 5. Conditionally merge the appendix document
        if (model.IncludeAppendix)
        {
            var appendixToInsert = new Document(appendixPath);
            var mergeBuilder = new DocumentBuilder(reportDoc);
            // Move to the bookmark location
            mergeBuilder.MoveToBookmark("AppendixHere");
            // Insert the appendix content at the bookmark position
            mergeBuilder.InsertDocument(appendixToInsert, ImportFormatMode.KeepSourceFormatting);
        }

        // 6. Build the report using LINQ Reporting Engine
        var engine = new ReportingEngine();
        engine.BuildReport(reportDoc, model, "model");

        // 7. Save the final report
        reportDoc.Save(resultPath);
    }
}

// Data model used by the LINQ Reporting template
public class ReportModel
{
    public string Title { get; set; } = string.Empty;
    public bool IncludeAppendix { get; set; }
    public string AppendixPlaceholder { get; set; } = string.Empty;
}
