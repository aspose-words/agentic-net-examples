using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define file names in the current working directory.
        string dictionaryPath = Path.Combine(Directory.GetCurrentDirectory(), "hyph_en_US.dic");
        string hyphenatedPdf = Path.Combine(Directory.GetCurrentDirectory(), "Hyphenated.pdf");
        string nonHyphenatedPdf = Path.Combine(Directory.GetCurrentDirectory(), "NonHyphenated.pdf");

        // Create a minimal hyphenation dictionary for English (US).
        // First line is the encoding, subsequent lines are word=hyphenated-pieces.
        File.WriteAllText(dictionaryPath,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Register the dictionary so Aspose.Words can hyphenate words in this language.
        Aspose.Words.Hyphenation.RegisterDictionary("en-US", dictionaryPath);

        // -----------------------------------------------------------------
        // Hyphenated document
        // -----------------------------------------------------------------
        Document hyphenatedDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(hyphenatedDoc);

        // Narrow page width forces line wrapping, making hyphenation visible.
        hyphenatedDoc.FirstSection.PageSetup.PageWidth = 200;   // points
        hyphenatedDoc.FirstSection.PageSetup.LeftMargin = 20;
        hyphenatedDoc.FirstSection.PageSetup.RightMargin = 20;

        // Write a paragraph containing words that have hyphenation patterns defined above.
        builder.Writeln("extraordinarycharacteristically communication");

        // Enable automatic hyphenation.
        hyphenatedDoc.HyphenationOptions.AutoHyphenation = true;

        // Save the hyphenated version as PDF.
        hyphenatedDoc.Save(hyphenatedPdf);
        if (!File.Exists(hyphenatedPdf))
            throw new InvalidOperationException("Hyphenated PDF was not created.");

        // -----------------------------------------------------------------
        // Non‑hyphenated document
        // -----------------------------------------------------------------
        Document nonHyphenatedDoc = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(nonHyphenatedDoc);

        nonHyphenatedDoc.FirstSection.PageSetup.PageWidth = 200;
        nonHyphenatedDoc.FirstSection.PageSetup.LeftMargin = 20;
        nonHyphenatedDoc.FirstSection.PageSetup.RightMargin = 20;

        builder2.Writeln("extraordinarycharacteristically communication");
        // AutoHyphenation remains false (default).

        // Save the non‑hyphenated version as PDF.
        nonHyphenatedDoc.Save(nonHyphenatedPdf);
        if (!File.Exists(nonHyphenatedPdf))
            throw new InvalidOperationException("Non‑hyphenated PDF was not created.");

        // -----------------------------------------------------------------
        // Compare file sizes
        // -----------------------------------------------------------------
        long hyphenatedSize = new FileInfo(hyphenatedPdf).Length;
        long nonHyphenatedSize = new FileInfo(nonHyphenatedPdf).Length;

        Console.WriteLine($"Hyphenated PDF size: {hyphenatedSize} bytes");
        Console.WriteLine($"Non‑hyphenated PDF size: {nonHyphenatedSize} bytes");
        Console.WriteLine($"Size difference: {Math.Abs(hyphenatedSize - nonHyphenatedSize)} bytes");
    }
}
