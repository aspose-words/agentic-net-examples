using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some Aspose.Words features)
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        const string templatePath = "Template.docx";
        const string outputPath = "ReportWithToc.docx";

        // -------------------------------------------------
        // 1. Create the template document programmatically
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a Table of Contents field at the beginning of the document.
        // The switches configure the TOC to include heading levels 1‑3 and make entries hyperlinked.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
        builder.InsertBreak(BreakType.PageBreak);

        // Add some headings so the TOC has entries.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.1");
        builder.Writeln("Section 1.2");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 2.1");

        // Save the template to disk (required before reporting)
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // 2. Load the template and build the report
        // -------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // Use the LINQ Reporting engine. No data source is needed for this scenario.
        ReportingEngine engine = new ReportingEngine();

        // Build the report – there are no LINQ tags in the template, but we still invoke the engine
        // to follow the required lifecycle pattern.
        engine.BuildReport(reportDoc, new object());

        // Update fields so the TOC is populated with the headings.
        reportDoc.UpdateFields();

        // -------------------------------------------------
        // 3. Save the final document
        // -------------------------------------------------
        reportDoc.Save(outputPath);
    }
}
