using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Define custom paragraph styles that we want to include in the TOC.
        Style quoteStyle = doc.Styles.Add(StyleType.Paragraph, "Quote");
        quoteStyle.Font.Size = 14;
        quoteStyle.Font.Color = Color.Blue;

        Style intenseQuoteStyle = doc.Styles.Add(StyleType.Paragraph, "Intense Quote");
        intenseQuoteStyle.Font.Size = 14;
        intenseQuoteStyle.Font.Color = Color.Red;
        intenseQuoteStyle.Font.Bold = true;

        // Insert a TOC that includes only the custom styles using the \t switch.
        // The list after \t is a comma‑separated list of "StyleName;Level" pairs.
        builder.InsertTableOfContents("\\t \"Quote;6,Intense Quote;7\" \\o \"1-3\" \\h \\z \\u");
        builder.InsertBreak(BreakType.PageBreak);

        // Add various paragraphs. Only those with the custom styles will appear in the TOC.

        // A built‑in heading style (not included in the \t list).
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Heading 1 – Not in TOC");

        // Paragraph with the "Quote" style (will be listed).
        builder.ParagraphFormat.Style = quoteStyle;
        builder.Writeln("Quote entry – Appears in TOC");

        // Normal paragraph (not included).
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Normal paragraph – Not in TOC");

        // Paragraph with the "Intense Quote" style (will be listed).
        builder.ParagraphFormat.Style = intenseQuoteStyle;
        builder.Writeln("Intense Quote entry – Appears in TOC");

        // Another built‑in heading style (not included).
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Heading 2 – Not in TOC");

        // Update all fields (including the TOC) to reflect the current document content.
        doc.UpdateFields();

        // Save the document to the current working directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TOC_CustomStyles.docx");
        doc.Save(outputPath);
    }
}
