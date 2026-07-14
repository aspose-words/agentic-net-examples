using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Paths for the temporary template and final report.
        string templatePath = "Template.docx";
        string reportPath = "Report.docx";

        // -------------------------------------------------
        // 1. Create the template document programmatically.
        // -------------------------------------------------
        Document templateDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(templateDoc);

        // Insert a Table of Contents that will pick up headings 1‑3.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
        builder.InsertBreak(BreakType.PageBreak);

        // Add sample headings that the TOC will reference.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1 – Introduction");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.1 – Background");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.2 – Scope");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2 – Details");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 2.1 – Analysis");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
        builder.Writeln("Subsection 2.1.1 – Data");

        // Save the template to disk.
        templateDoc.Save(templatePath);

        // -------------------------------------------------
        // 2. Load the template and run the LINQ Reporting engine.
        // -------------------------------------------------
        Document reportDoc = new Document(templatePath);

        // The template does not contain any LINQ Reporting tags,
        // so we can pass an empty data source.
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(reportDoc, new object());

        // Update fields (TOC) to reflect the headings.
        reportDoc.UpdateFields();

        // -------------------------------------------------
        // 3. Save the final report.
        // -------------------------------------------------
        reportDoc.Save(reportPath);
    }
}
