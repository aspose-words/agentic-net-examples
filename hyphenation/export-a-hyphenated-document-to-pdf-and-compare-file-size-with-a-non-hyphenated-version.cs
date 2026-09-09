using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class HyphenationComparison
{
    public static void Main()
    {
        // Prepare a minimal hyphenation dictionary for English (US).
        const string dictPath = "hyph_en_US.dic";
        File.WriteAllText(dictPath,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Register the dictionary so Aspose.Words can hyphenate the words.
        Hyphenation.RegisterDictionary("en-US", dictPath);
        if (!Hyphenation.IsDictionaryRegistered("en-US"))
            throw new InvalidOperationException("Failed to register the hyphenation dictionary.");

        // Create a document with long words that require hyphenation.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Size = 24;
        builder.Writeln("extraordinarycharacteristically internationalization communication");
        // Narrow the page width to force line wrapping and hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 200; // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // ---------- Hyphenated PDF ----------
        doc.HyphenationOptions.AutoHyphenation = true;
        const string hyphenatedPdf = "hyphenated.pdf";
        doc.Save(hyphenatedPdf);
        if (!File.Exists(hyphenatedPdf))
            throw new InvalidOperationException("Hyphenated PDF was not created.");

        // ---------- Non‑hyphenated PDF ----------
        doc.HyphenationOptions.AutoHyphenation = false;
        const string nonHyphenatedPdf = "nonhyphenated.pdf";
        doc.Save(nonHyphenatedPdf);
        if (!File.Exists(nonHyphenatedPdf))
            throw new InvalidOperationException("Non‑hyphenated PDF was not created.");

        // Compare file sizes.
        long hyphenatedSize = new FileInfo(hyphenatedPdf).Length;
        long nonHyphenatedSize = new FileInfo(nonHyphenatedPdf).Length;

        Console.WriteLine($"Hyphenated PDF size: {hyphenatedSize} bytes");
        Console.WriteLine($"Non‑hyphenated PDF size: {nonHyphenatedSize} bytes");
        Console.WriteLine($"Size difference (non‑hyphenated - hyphenated): {nonHyphenatedSize - hyphenatedSize} bytes");
    }
}
