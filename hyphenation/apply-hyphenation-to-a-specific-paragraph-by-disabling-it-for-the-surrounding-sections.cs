using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        var doc = new Document();

        // Use DocumentBuilder to add content.
        var builder = new DocumentBuilder(doc);

        // Narrow the page width to force line wrapping and make hyphenation visible.
        doc.FirstSection.PageSetup.PageWidth = 300; // points (~4.2 inches)
        doc.FirstSection.PageSetup.PageHeight = 842; // A4 height.

        // Enable automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Set a valid hyphenation zone (default is 360 = 0.25 inch).
        // Setting it to 0 throws an ArgumentOutOfRangeException, so we keep the default.
        doc.HyphenationOptions.HyphenationZone = 360;

        // ---------- First paragraph (hyphenation suppressed) ----------
        builder.Writeln("This is a long paragraph that will demonstrate hyphenation. It contains many words that could be split across lines if hyphenation is enabled.");
        var firstParagraph = doc.FirstSection.Body.Paragraphs[0];
        firstParagraph.ParagraphFormat.SuppressAutoHyphens = true; // Disable hyphenation.

        // ---------- Target paragraph (hyphenation enabled) ----------
        builder.Writeln("Another long paragraph that should be hyphenated when the line breaks occur. The quick brown fox jumps over the lazy dog repeatedly to fill the line.");

        // ---------- Third paragraph (hyphenation suppressed) ----------
        builder.Writeln("Final paragraph that will not be hyphenated despite line length, demonstrating suppression.");
        var thirdParagraph = doc.FirstSection.Body.Paragraphs[2];
        thirdParagraph.ParagraphFormat.SuppressAutoHyphens = true; // Disable hyphenation.

        // Save the document to a PDF file in the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "HyphenationExample.pdf");
        doc.Save(outputPath);
    }
}
