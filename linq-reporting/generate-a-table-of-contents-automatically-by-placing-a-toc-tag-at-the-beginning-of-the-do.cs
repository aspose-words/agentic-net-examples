using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;
using Aspose.Words.Fields;

public class ReportModel
{
    // Dummy property to satisfy non‑nullable requirement.
    public string Dummy { get; set; } = "placeholder";
}

public class Program
{
    public static void Main()
    {
        // Register code page provider for Aspose.Words (required for some encodings).
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        // Paths for the template and the final report.
        string templatePath = "Template.docx";
        string outputPath = "ReportWithTOC.docx";

        // -------------------------------------------------
        // 1. Create the template document programmatically.
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a Table of Contents field at the very beginning.
        // Switches: include heading levels 1‑3, add hyperlinks, hide page numbers in web view, use outline levels.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");

        // Start a new page for the actual content.
        builder.InsertBreak(BreakType.PageBreak);

        // Add sample headings that the TOC will pick up.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1: Introduction");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.1: Background");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.2: Objectives");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2: Details");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 2.1: Implementation");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
        builder.Writeln("Subsection 2.1.1: Code Overview");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // 2. Load the template and build the report.
        // -------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // Create a dummy data source; the template does not reference it,
        // but the ReportingEngine requires a root object.
        ReportModel model = new ReportModel();

        // Initialize the reporting engine.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;

        // Build the report. The root name "model" matches the object passed.
        engine.BuildReport(reportDoc, model, "model");

        // Update all fields (including the TOC) so that the table of contents reflects the headings.
        reportDoc.UpdateFields();

        // Save the final document.
        reportDoc.Save(outputPath);
    }
}
