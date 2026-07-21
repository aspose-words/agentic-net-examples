using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Register code page provider (required for some encodings)
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        // -------------------- Create template --------------------
        Document template = new Document();
        DocumentBuilder builder = new DocumentBuilder(template);

        // Insert a Table of Contents that picks up headings 1‑3 and creates hyperlinks
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
        builder.InsertBreak(BreakType.PageBreak);

        // Add sample headings – these will become TOC entries
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.1");
        builder.Writeln("Section 1.2");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
        builder.Writeln("Subsection 2.1.1");

        // Add a simple data list using LINQ Reporting tags
        builder.Writeln("\nItems:");
        builder.Writeln("<<foreach [item in Items]>>");
        builder.Writeln("- <<[item]>>");
        builder.Writeln("<</foreach>>");

        // Save the template to disk (required before building the report)
        string templatePath = Path.Combine(Environment.CurrentDirectory, "Template.docx");
        template.Save(templatePath);

        // -------------------- Prepare data model --------------------
        var model = new ReportModel
        {
            Items = new List<string> { "Apple", "Banana", "Cherry" }
        };

        // -------------------- Build report --------------------
        Document report = new Document(templatePath);
        ReportingEngine engine = new ReportingEngine();
        engine.BuildReport(report, model, "model");

        // Update fields so the TOC reflects the generated headings
        report.UpdateFields();

        // Save the final document
        string outputPath = Path.Combine(Environment.CurrentDirectory, "ReportWithTOC.docx");
        report.Save(outputPath);
    }

    // Public data model used by the LINQ Reporting engine
    public class ReportModel
    {
        public List<string> Items { get; set; } = new();
    }
}
