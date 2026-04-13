using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "HyphenationReport.pdf");

        // Create a new blank document.
        Document doc = new Document();

        // Make the page narrow so that words are forced to wrap and hyphenate.
        Section section = doc.FirstSection;
        section.PageSetup.PageWidth = 300;   // 300 points (~4.2 cm)
        section.PageSetup.PageHeight = 842; // Standard A4 height.

        // Enable automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;
        // Optional: configure hyphenation limits.
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenationZone = 360; // 0.25 inch.

        DocumentBuilder builder = new DocumentBuilder(doc);

        // ----- Heading (hyphenation disabled) -----
        // Apply a heading style.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        // Suppress hyphenation for this paragraph.
        builder.ParagraphFormat.SuppressAutoHyphens = true;
        builder.Writeln("This is a very long heading that would normally be hyphenated if it wrapped, but hyphenation is suppressed for headings.");

        // ----- Body paragraph (hyphenation enabled) -----
        // Reset the suppression flag so hyphenation works for following paragraphs.
        builder.ParagraphFormat.SuppressAutoHyphens = false;
        // Use normal style for body text.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Font.Size = 12;
        // Add a paragraph with long words to trigger hyphenation.
        builder.Writeln(
            "This body paragraph contains words such as antidisestablishmentarianism and pneumonoultramicroscopicsilicovolcanoconiosis that are long enough to be split across lines. " +
            "When the line reaches the right margin, Aspose.Words will automatically insert hyphens according to the language dictionary.");

        // Save the document as PDF to render the layout and hyphenation.
        doc.Save(outputPath, SaveFormat.Pdf);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the output file at '{outputPath}'.");

        // The example finishes without requiring user interaction.
    }
}
