using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Paths for the template and the generated report.
        string templatePath = Path.Combine(Environment.CurrentDirectory, "Template.docx");
        string reportPath = Path.Combine(Environment.CurrentDirectory, "Report.docx");

        // -----------------------------------------------------------------
        // 1. Create the template document programmatically.
        // -----------------------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a Table of Contents field at the beginning of the document.
        // The switches configure the TOC to include heading levels 1‑3, add hyperlinks, etc.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
        builder.InsertBreak(BreakType.PageBreak);

        // Add sample content with heading styles that the TOC will pick up.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.1");
        builder.Writeln("Section 1.2");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 2.1");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
        builder.Writeln("Subsection 2.1.1");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -----------------------------------------------------------------
        // 2. Load the template and build the report using LINQ Reporting.
        // -----------------------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // The template does not contain any data‑binding tags, but we still use the
        // ReportingEngine to demonstrate the required workflow.
        ReportingEngine engine = new ReportingEngine();
        engine.Options = ReportBuildOptions.None;

        // BuildReport with an empty data source (no root object needed).
        engine.BuildReport(reportDoc, new object());

        // Update fields so that the TOC reflects the headings added above.
        reportDoc.UpdateFields();

        // Save the final document.
        reportDoc.Save(reportPath);
    }
}
