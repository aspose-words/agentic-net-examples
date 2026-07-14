using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class HyphenationPdfComparison
{
    public static void Main()
    {
        // Output file names
        const string dictPath = "hyph_en_US.dic";
        const string hyphenatedPdf = "hyphenated.pdf";
        const string nonHyphenatedPdf = "nonhyphenated.pdf";

        // Clean up any previous artifacts
        if (File.Exists(dictPath)) File.Delete(dictPath);
        if (File.Exists(hyphenatedPdf)) File.Delete(hyphenatedPdf);
        if (File.Exists(nonHyphenatedPdf)) File.Delete(nonHyphenatedPdf);

        // Create a minimal hyphenation dictionary for English (US)
        // Format: first line "UTF-8", then word=hyphenation-points
        File.WriteAllText(dictPath,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Register the dictionary for the "en-US" locale
        Hyphenation.RegisterDictionary("en-US", dictPath);

        // Build a document containing long words that can be hyphenated
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow the page width to force line wrapping
        doc.FirstSection.PageSetup.PageWidth = 200; // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Sample text with words present in the dictionary
        builder.Font.Size = 12;
        builder.Writeln("extraordinarycharacteristically internationalization communication");
        builder.Writeln("extraordinarycharacteristically internationalization communication");

        // ------------------- Hyphenated version -------------------
        doc.HyphenationOptions.AutoHyphenation = true;
        // Use the default hyphenation zone (360 = 0.25 inch) – this value is valid
        doc.HyphenationOptions.HyphenationZone = 360;
        doc.Save(hyphenatedPdf, SaveFormat.Pdf);

        if (!File.Exists(hyphenatedPdf))
            throw new InvalidOperationException("Hyphenated PDF was not created.");

        // ------------------- Non‑hyphenated version -------------------
        // Disable automatic hyphenation and re‑save
        doc.HyphenationOptions.AutoHyphenation = false;
        doc.Save(nonHyphenatedPdf, SaveFormat.Pdf);

        if (!File.Exists(nonHyphenatedPdf))
            throw new InvalidOperationException("Non‑hyphenated PDF was not created.");

        // Compare file sizes
        long hyphenatedSize = new FileInfo(hyphenatedPdf).Length;
        long nonHyphenatedSize = new FileInfo(nonHyphenatedPdf).Length;
        long difference = Math.Abs(hyphenatedSize - nonHyphenatedSize);

        Console.WriteLine($"Hyphenated PDF size: {hyphenatedSize} bytes");
        Console.WriteLine($"Non‑hyphenated PDF size: {nonHyphenatedSize} bytes");
        Console.WriteLine($"Size difference: {difference} bytes");
    }
}
