using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Create a minimal hyphenation dictionary for English (US).
        const string dictPath = "hyph_en_US.dic";
        File.WriteAllText(dictPath,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Register the dictionary so that hyphenation can be applied.
        Hyphenation.RegisterDictionary("en-US", dictPath);

        // Build a document containing long words that can be hyphenated.
        Document hyphenDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(hyphenDoc);
        builder.Writeln("extraordinarycharacteristically internationalization communication");

        // Narrow the page width to force line wrapping.
        hyphenDoc.FirstSection.PageSetup.PageWidth = 200;
        hyphenDoc.FirstSection.PageSetup.LeftMargin = 20;
        hyphenDoc.FirstSection.PageSetup.RightMargin = 20;

        // Enable automatic hyphenation.
        hyphenDoc.HyphenationOptions.AutoHyphenation = true;

        const string hyphenPdf = "Hyphenated.pdf";
        hyphenDoc.Save(hyphenPdf, SaveFormat.Pdf);
        if (!File.Exists(hyphenPdf))
            throw new InvalidOperationException("Hyphenated PDF was not created.");

        // Build the same document without hyphenation.
        Document nonHyphenDoc = new Document();
        DocumentBuilder nbuilder = new DocumentBuilder(nonHyphenDoc);
        nbuilder.Writeln("extraordinarycharacteristically internationalization communication");
        nonHyphenDoc.FirstSection.PageSetup.PageWidth = 200;
        nonHyphenDoc.FirstSection.PageSetup.LeftMargin = 20;
        nonHyphenDoc.FirstSection.PageSetup.RightMargin = 20;
        // Ensure hyphenation is disabled (default, but set explicitly for clarity).
        nonHyphenDoc.HyphenationOptions.AutoHyphenation = false;

        const string nonHyphenPdf = "NonHyphenated.pdf";
        nonHyphenDoc.Save(nonHyphenPdf, SaveFormat.Pdf);
        if (!File.Exists(nonHyphenPdf))
            throw new InvalidOperationException("Non‑hyphenated PDF was not created.");

        // Compare file sizes.
        long hyphenSize = new FileInfo(hyphenPdf).Length;
        long nonHyphenSize = new FileInfo(nonHyphenPdf).Length;

        Console.WriteLine($"Hyphenated PDF size: {hyphenSize} bytes");
        Console.WriteLine($"Non‑hyphenated PDF size: {nonHyphenSize} bytes");
        Console.WriteLine($"Size difference: {Math.Abs(hyphenSize - nonHyphenSize)} bytes");
    }
}
