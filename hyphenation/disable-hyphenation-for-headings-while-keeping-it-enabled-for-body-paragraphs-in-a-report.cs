using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Path for the temporary hyphenation dictionary.
        const string dictPath = "hyph_en_US.dic";

        // Create a minimal English hyphenation dictionary.
        // The format is: first line "UTF-8", then word=hyphenated-pieces per line.
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

        // Narrow the page width so that long words need to wrap and hyphenate.
        // Width is in points (1 point = 1/72 inch). 200 points ≈ 2.78 inches.
        doc.FirstSection.PageSetup.PageWidth = 200;
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Enable automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenationZone = 360; // 0.25 inch

        // ---------- Headings (hyphenation disabled) ----------
        // Heading 1
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.ParagraphFormat.SuppressAutoHyphens = true; // Disable hyphenation for this paragraph.
        builder.Writeln("Heading 1: This is a very long heading that could be hyphenated but we suppress it.");

        // Heading 2
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.ParagraphFormat.SuppressAutoHyphens = true;
        builder.Writeln("Heading 2: Another lengthy heading that should stay on one line without hyphens.");

        // ---------- Body paragraphs (hyphenation enabled) ----------
        // Reset to normal style and enable hyphenation.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.ParagraphFormat.SuppressAutoHyphens = false;

        // Body paragraph containing words that can be hyphenated.
        builder.Writeln(
            "Body paragraph: The word extraordinarycharacteristically demonstrates how hyphenation works. " +
            "Another example is internationalization which may also be split across lines. " +
            "Communication between components often requires clear formatting.");

        // Add a second body paragraph to ensure multiple lines.
        builder.Writeln(
            "Additional body text: Aspose.Words provides powerful APIs for document generation, " +
            "including automatic hyphenation, style management, and layout control.");

        // Save the document to PDF.
        const string outputPath = "HyphenationReport.pdf";
        doc.Save(outputPath, SaveFormat.Pdf);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The expected PDF output was not created.");

        // Clean up the temporary dictionary file.
        if (File.Exists(dictPath))
            File.Delete(dictPath);
    }
}
