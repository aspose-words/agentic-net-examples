using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Folder where the output document will be saved.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a new blank document and a builder to populate it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a TOC that includes only the built‑in heading styles (1‑3) and two custom styles:
        // "Quote" (TOC level 6) and "Intense Quote" (TOC level 7).
        // The \t switch uses a comma‑separated list of style name and level pairs.
        builder.InsertTableOfContents("\\t \"Quote;6,Intense Quote;7\" \\o \"1-3\" \\h \\z \\u");
        builder.InsertBreak(BreakType.PageBreak);

        // Create custom styles that will be referenced by the TOC.
        Style quoteStyle = doc.Styles.Add(StyleType.Paragraph, "Quote");
        quoteStyle.Font.Italic = true;

        Style intenseQuoteStyle = doc.Styles.Add(StyleType.Paragraph, "Intense Quote");
        intenseQuoteStyle.Font.Bold = true;
        intenseQuoteStyle.Font.Color = System.Drawing.Color.DarkRed;

        // Populate the document with headings and custom‑styled paragraphs.
        InsertNewPageWithHeading(builder, "Heading 1", StyleIdentifier.Heading1);
        InsertNewPageWithHeading(builder, "Heading 1.1", StyleIdentifier.Heading2);
        InsertNewPageWithHeading(builder, "Quote paragraph", "Quote");
        InsertNewPageWithHeading(builder, "Intense Quote paragraph", "Intense Quote");
        InsertNewPageWithHeading(builder, "Heading 2", StyleIdentifier.Heading1);
        InsertNewPageWithHeading(builder, "Heading 2.1", StyleIdentifier.Heading2);

        // Update all fields (the TOC) so that it reflects the current content.
        doc.UpdateFields();

        // Save the document.
        string outputPath = Path.Combine(artifactsDir, "TocSpecificStyles.docx");
        doc.Save(outputPath);
    }

    // Helper method that starts a new page and writes a paragraph using the specified style.
    private static void InsertNewPageWithHeading(DocumentBuilder builder, string text, StyleIdentifier styleId)
    {
        builder.InsertBreak(BreakType.PageBreak);
        builder.ParagraphFormat.StyleIdentifier = styleId;
        builder.Writeln(text);
    }

    // Overload for custom style names.
    private static void InsertNewPageWithHeading(DocumentBuilder builder, string text, string styleName)
    {
        builder.InsertBreak(BreakType.PageBreak);
        builder.ParagraphFormat.Style = builder.Document.Styles[styleName];
        builder.Writeln(text);
    }
}
