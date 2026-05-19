using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "TocWithCustomStyles.docx");

        // Add custom paragraph styles that we want to appear in the TOC.
        Style quoteStyle = doc.Styles.Add(StyleType.Paragraph, "Quote");
        quoteStyle.Font.Italic = true;

        Style intenseQuoteStyle = doc.Styles.Add(StyleType.Paragraph, "Intense Quote");
        intenseQuoteStyle.Font.Bold = true;
        intenseQuoteStyle.Font.Italic = true;

        // Insert a TOC that includes only the custom styles using the \t switch.
        // The syntax is: \t "StyleName;Level,StyleName;Level"
        // Here we map "Quote" to TOC level 1 and "Intense Quote" to TOC level 2.
        builder.InsertTableOfContents("\\t \"Quote;1,Intense Quote;2\" \\h \\z \\u");
        builder.InsertBreak(BreakType.PageBreak);

        // Add content using the custom styles.
        builder.ParagraphFormat.Style = quoteStyle;
        builder.Writeln("First quoted paragraph.");

        builder.ParagraphFormat.Style = intenseQuoteStyle;
        builder.Writeln("First intense quoted paragraph.");

        // Add some regular headings to show they are NOT included in the TOC.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Regular Heading 1");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Regular Heading 2");

        // Add more custom styled paragraphs on a new page.
        builder.InsertBreak(BreakType.PageBreak);
        builder.ParagraphFormat.Style = quoteStyle;
        builder.Writeln("Second quoted paragraph.");

        builder.ParagraphFormat.Style = intenseQuoteStyle;
        builder.Writeln("Second intense quoted paragraph.");

        // Update fields so the TOC reflects the current document structure.
        doc.UpdateFields();

        // Save the document.
        doc.Save(outputPath);
    }
}
