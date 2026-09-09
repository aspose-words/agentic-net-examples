using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Paths for the dictionary and the output PDF.
        const string dictionaryPath = "hyph_en_US.dic";
        const string outputPath = "Report.pdf";

        // Create a minimal hyphenation dictionary for English (US).
        // The format: first line is the encoding, subsequent lines are word=hyphenation-points.
        File.WriteAllText(dictionaryPath,
@"UTF-8
extraordinarycharacteristically=ex-tra-or-di-nary-char-ac-ter-is-ti-cal-ly
internationalization=in-ter-na-tion-al-i-za-tion
communication=com-mu-ni-ca-tion");

        // Register the dictionary so that Aspose.Words can hyphenate English text.
        Hyphenation.RegisterDictionary("en-US", dictionaryPath);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Enable automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;
        // Optional: tweak hyphenation settings.
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch

        // Narrow the page width to force line wrapping and hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 300; // points (~4.17 inches)
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // ---------- Heading (hyphenation disabled) ----------
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        // Suppress hyphenation for this heading.
        builder.ParagraphFormat.SuppressAutoHyphens = true;
        builder.Font.Size = 24;
        builder.Writeln("Heading: extraordinarycharacteristically internationalization communication");

        // Add a blank line between heading and body.
        builder.Writeln();

        // ---------- Body paragraph (hyphenation enabled) ----------
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        // Ensure hyphenation is allowed for body paragraphs.
        builder.ParagraphFormat.SuppressAutoHyphens = false;
        builder.Font.Size = 12;
        builder.Writeln(
            "Body: The quick brown fox jumps over the lazy dog. " +
            "This paragraph contains the word extraordinarycharacteristically which is long enough to be hyphenated " +
            "when it reaches the end of the line. The same applies to internationalization and communication.");

        // Save the document to PDF.
        doc.Save(outputPath, SaveFormat.Pdf);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Expected output file '{outputPath}' was not created.");

        // Clean up the temporary dictionary file.
        if (File.Exists(dictionaryPath))
            File.Delete(dictionaryPath);
    }
}
