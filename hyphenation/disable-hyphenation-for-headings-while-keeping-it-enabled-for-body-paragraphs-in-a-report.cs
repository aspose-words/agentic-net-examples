using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class HyphenationReport
{
    public static void Main()
    {
        // Path for the temporary hyphenation dictionary.
        const string dictPath = "hyph_en_US.dic";

        // Create a minimal English hyphenation dictionary.
        // The first line must specify the encoding.
        // Subsequent lines define hyphenation patterns for words used in the document.
        File.WriteAllText(dictPath,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Register the dictionary for the "en-US" locale.
        Hyphenation.RegisterDictionary("en-US", dictPath);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Enable automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;
        // Optional: tighten hyphenation zone to make hyphenation more likely.
        doc.HyphenationOptions.HyphenationZone = 360; // 0.25 inch

        // Narrow the page width so that long words need to wrap.
        doc.FirstSection.PageSetup.PageWidth = 300; // points (~4.2 inches)
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // ---------- Heading (hyphenation disabled) ----------
        builder.Font.Size = 24;
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        // Suppress hyphenation for this heading.
        builder.ParagraphFormat.SuppressAutoHyphens = true;
        builder.Writeln("Heading: extraordinarycharacteristically");

        // Reset paragraph formatting for body text.
        builder.Font.Size = 12;
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.ParagraphFormat.SuppressAutoHyphens = false; // Enable hyphenation.

        // ---------- Body paragraph (hyphenation enabled) ----------
        builder.Writeln(
            "The body of the report contains long words that will be hyphenated when needed: " +
            "extraordinarycharacteristically internationalization communication extraordinarycharacteristically.");

        // Save the document to PDF (any format works; PDF shows layout).
        const string outputPath = "HyphenationReport.pdf";
        doc.Save(outputPath, SaveFormat.Pdf);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The expected output file was not created.");

        // Clean up the temporary dictionary file.
        File.Delete(dictPath);
    }
}
