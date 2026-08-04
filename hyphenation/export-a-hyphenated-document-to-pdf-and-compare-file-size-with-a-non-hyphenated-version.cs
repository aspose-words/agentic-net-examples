using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Paths for temporary files.
        const string dictionaryPath = "hyph_en_US.dic";
        const string nonHyphenatedPdf = "nonhyphenated.pdf";
        const string hyphenatedPdf = "hyphenated.pdf";

        // Create a minimal hyphenation dictionary for English (US).
        // The first line must specify the encoding, followed by word‑hyphenation patterns.
        File.WriteAllText(dictionaryPath,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Register the dictionary so that Aspose.Words can hyphenate English text.
        Hyphenation.RegisterDictionary("en-US", dictionaryPath);

        // -----------------------------------------------------------------
        // Create the base document with sample text that can be hyphenated.
        // -----------------------------------------------------------------
        Document baseDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(baseDoc);

        // Use a relatively large font to make line wrapping more likely.
        builder.Font.Size = 24;
        builder.Writeln(
            "extraordinarycharacteristically internationalization communication " +
            "extraordinarycharacteristically internationalization communication");

        // Narrow the page width to force words onto new lines where hyphenation may occur.
        baseDoc.FirstSection.PageSetup.PageWidth = 200;
        baseDoc.FirstSection.PageSetup.LeftMargin = 20;
        baseDoc.FirstSection.PageSetup.RightMargin = 20;

        // ---------------------------------------------------------------
        // Save the document without automatic hyphenation (default state).
        // ---------------------------------------------------------------
        baseDoc.Save(nonHyphenatedPdf);
        if (!File.Exists(nonHyphenatedPdf))
            throw new InvalidOperationException("Non‑hyphenated PDF was not created.");

        // ---------------------------------------------------------------
        // Clone the base document, enable automatic hyphenation, and save.
        // ---------------------------------------------------------------
        Document hyphenatedDoc = (Document)baseDoc.Clone(true);
        hyphenatedDoc.HyphenationOptions.AutoHyphenation = true;
        hyphenatedDoc.Save(hyphenatedPdf);
        if (!File.Exists(hyphenatedPdf))
            throw new InvalidOperationException("Hyphenated PDF was not created.");

        // ---------------------------------------------------------------
        // Compare file sizes.
        // ---------------------------------------------------------------
        long sizeNon = new FileInfo(nonHyphenatedPdf).Length;
        long sizeHy = new FileInfo(hyphenatedPdf).Length;

        Console.WriteLine($"Non‑hyphenated PDF size: {sizeNon} bytes");
        Console.WriteLine($"Hyphenated PDF size: {sizeHy} bytes");
        Console.WriteLine($"Size difference: {sizeHy - sizeNon} bytes");
    }
}
