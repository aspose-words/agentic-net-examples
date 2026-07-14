using System;
using System.IO;
using Aspose.Words;

public class HyphenationRangeExample
{
    public static void Main()
    {
        // Path for the temporary hyphenation dictionary.
        const string dictPath = "hyph_en_US.dic";

        // Create a minimal hyphenation dictionary for English (US).
        // The dictionary follows the OpenOffice hyphenation pattern.
        File.WriteAllText(dictPath,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Register the dictionary so that Aspose.Words can hyphenate English text.
        Hyphenation.RegisterDictionary("en-US", dictPath);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow the page width to force line wrapping and make hyphenation visible.
        doc.FirstSection.PageSetup.PageWidth = 200; // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Enable automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;
        // HyphenationZone must be a non‑negative value; use the default (360) which equals 0.25 inch.
        doc.HyphenationOptions.HyphenationZone = 360;
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenateCaps = true;

        // First paragraph – hyphenation is allowed (default).
        builder.Writeln("extraordinarycharacteristically internationalization communication");

        // Second paragraph – explicitly suppress hyphenation for this range.
        builder.ParagraphFormat.SuppressAutoHyphens = true;
        builder.Writeln("extraordinarycharacteristically internationalization communication");

        // Save the document to PDF to render the layout.
        const string outputPath = "HyphenatedRange.pdf";
        doc.Save(outputPath, SaveFormat.Pdf);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Expected output file '{outputPath}' was not created.");

        // Clean up the temporary dictionary file.
        if (File.Exists(dictPath))
            File.Delete(dictPath);
    }
}
