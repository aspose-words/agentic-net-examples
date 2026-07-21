using System;
using Aspose.Words;
using Aspose.Words.Reporting;

public class Program
{
    public static void Main()
    {
        // Sample data model with headings.
        var model = new ReportModel
        {
            Title1 = "Chapter 1: Introduction",
            Title2 = "Section 1.1: Overview",
            Title3 = "Subsection 1.1.1: Details"
        };

        // Create a blank document and a builder.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Insert a Table of Contents field that will pick up headings 1‑3.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
        builder.InsertBreak(BreakType.PageBreak);

        // Heading level 1 with an expression tag.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("<<[model.Title1]>>");

        // Heading level 2 with an expression tag.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("<<[model.Title2]>>");

        // Heading level 3 with an expression tag.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
        builder.Writeln("<<[model.Title3]>>");

        // Populate the template using the LINQ Reporting engine.
        var engine = new ReportingEngine();
        engine.BuildReport(doc, model, "model");

        // Update fields so the TOC reflects the generated headings.
        doc.UpdateFields();

        // Save the resulting document.
        doc.Save("Output.docx");
    }

    // Public data model used by the template.
    public class ReportModel
    {
        public string Title1 { get; set; } = "";
        public string Title2 { get; set; } = "";
        public string Title3 { get; set; } = "";
    }
}
