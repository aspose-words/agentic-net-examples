using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Path for the temporary hyphenation dictionary.
        const string dictionaryPath = "hyph_en_US.dic";

        // Create a minimal hyphenation dictionary for English (US).
        // The first line must be the encoding, followed by word=hyphenation-pattern lines.
        File.WriteAllText(dictionaryPath,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Register the dictionary so that Aspose.Words can hyphenate English text.
        Hyphenation.RegisterDictionary("en-US", dictionaryPath);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow the page width to make hyphenation more likely to occur.
        doc.FirstSection.PageSetup.PageWidth = 300; // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Enable automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;
        // Optional: configure additional hyphenation options.
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch

        // ----- Heading (hyphenation disabled) -----
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        // Suppress hyphenation for this paragraph (heading).
        builder.ParagraphFormat.SuppressAutoHyphens = true;
        builder.Writeln("Heading: extraordinarycharacteristically internationalization communication");

        // ----- Body paragraph (hyphenation enabled) -----
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        // Ensure hyphenation is not suppressed for body text.
        builder.ParagraphFormat.SuppressAutoHyphens = false;
        builder.Writeln("Body: extraordinarycharacteristically internationalization communication " +
                        "are long words that can be split across lines when hyphenation is active.");

        // Save the document to PDF.
        const string outputPath = "Report.pdf";
        doc.Save(outputPath, SaveFormat.Pdf);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The expected output file was not created.");

        // Clean up the temporary dictionary file.
        File.Delete(dictionaryPath);
    }
}
