using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -----------------------------------------------------------------
        // Define custom styles that we want to appear in the Table of Contents.
        // The style name is followed by a semicolon and the TOC level.
        // Multiple styles are separated by commas.
        // Example switch: \t "Quote;6,Intense Quote;7"
        // -----------------------------------------------------------------
        Style quoteStyle = doc.Styles.Add(StyleType.Paragraph, "Quote");
        quoteStyle.Font.Italic = true;
        quoteStyle.Font.Color = System.Drawing.Color.DarkGreen;

        Style intenseQuoteStyle = doc.Styles.Add(StyleType.Paragraph, "Intense Quote");
        intenseQuoteStyle.Font.Italic = true;
        intenseQuoteStyle.Font.Bold = true;
        intenseQuoteStyle.Font.Color = System.Drawing.Color.DarkRed;

        // Insert a TOC field that includes only the built‑in Heading styles (1‑3)
        // and the two custom styles defined above.
        // The \t switch specifies the custom styles with their TOC levels.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u \\t \"Quote;6,Intense Quote;7\"");

        // Add a page break so that the TOC appears on its own page.
        builder.InsertBreak(BreakType.PageBreak);

        // -----------------------------------------------------------------
        // Populate the document with headings and custom styled paragraphs.
        // -----------------------------------------------------------------

        // Built‑in heading style – will appear in the TOC (level 1‑3).
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.1");
        builder.Writeln("Section 1.2");

        // Custom style "Quote" – will appear in the TOC at level 6.
        builder.ParagraphFormat.Style = quoteStyle;
        builder.Writeln("This is a quoted paragraph that should be listed in the TOC.");

        // Custom style "Intense Quote" – will appear in the TOC at level 7.
        builder.ParagraphFormat.Style = intenseQuoteStyle;
        builder.Writeln("This is an intense quoted paragraph that should also be listed.");

        // Another built‑in heading – will appear in the TOC.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2");

        // -----------------------------------------------------------------
        // Update fields to generate the TOC entries and save the document.
        // -----------------------------------------------------------------
        doc.UpdateFields();

        string outputPath = Path.Combine(Environment.CurrentDirectory, "TOC_Styles.docx");
        doc.Save(outputPath);
    }
}
