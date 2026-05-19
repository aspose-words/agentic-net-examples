using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class HyphenationRangeExample
{
    public static void Main()
    {
        // Create a minimal hyphenation dictionary for English (US).
        string dictPath = Path.Combine(Directory.GetCurrentDirectory(), "hyph_en_US.dic");
        File.WriteAllText(dictPath,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Register the dictionary so that hyphenation can be performed.
        Hyphenation.RegisterDictionary("en-US", dictPath);
        if (!Hyphenation.IsDictionaryRegistered("en-US"))
            throw new InvalidOperationException("Failed to register hyphenation dictionary.");

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow the page width to force line wrapping.
        doc.FirstSection.PageSetup.PageWidth = 300; // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Sample long text that can be hyphenated.
        string longText = "extraordinarycharacteristically internationalization communication";

        // First paragraph: hyphenation suppressed.
        builder.ParagraphFormat.SuppressAutoHyphens = true;
        builder.Writeln(longText);

        // Second paragraph: hyphenation allowed (selected range).
        builder.ParagraphFormat.SuppressAutoHyphens = false;
        builder.Writeln(longText);

        // Enable automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenationZone = 360; // default
        doc.HyphenationOptions.HyphenateCaps = true;

        // Save the document to PDF to materialize layout.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "HyphenatedRange.pdf");
        doc.Save(outputPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Expected PDF output was not created.");

        // Clean up temporary dictionary file.
        File.Delete(dictPath);
    }
}
