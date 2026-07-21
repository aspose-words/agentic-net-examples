using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using static Aspose.Words.Hyphenation; // Hyphenation is a static class, import its members

public class Program
{
    public static void Main()
    {
        // Create a minimal hyphenation dictionary locally.
        const string dictFile = "hyph_en_US.dic";

        if (!File.Exists(dictFile))
        {
            File.WriteAllText(dictFile,
                "UTF-8\n" +
                "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
                "internationalization=in-ter-na-tion-al-i-za-tion\n" +
                "communication=com-mu-ni-ca-tion\n");
        }

        // Register the dictionary for English (US).
        RegisterDictionary("en-US", dictFile);

        // Build two sample documents: one with hyphenation enabled, one without.
        Document hyphenatedDoc = CreateSampleDocument(autoHyphenation: true);
        Document nonHyphenatedDoc = CreateSampleDocument(autoHyphenation: false);

        // Measure PDF generation time with hyphenation.
        var sw = Stopwatch.StartNew();
        hyphenatedDoc.Save("Hyphenated.pdf", SaveFormat.Pdf);
        sw.Stop();
        long hyphenatedTimeMs = sw.ElapsedMilliseconds;

        // Measure PDF generation time without hyphenation.
        sw.Restart();
        nonHyphenatedDoc.Save("NonHyphenated.pdf", SaveFormat.Pdf);
        sw.Stop();
        long nonHyphenatedTimeMs = sw.ElapsedMilliseconds;

        // Validate that the PDFs were created.
        if (!File.Exists("Hyphenated.pdf"))
            throw new InvalidOperationException("Hyphenated.pdf was not created.");
        if (!File.Exists("NonHyphenated.pdf"))
            throw new InvalidOperationException("NonHyphenated.pdf was not created.");

        // Output timing results.
        Console.WriteLine($"PDF generation with hyphenation: {hyphenatedTimeMs} ms");
        Console.WriteLine($"PDF generation without hyphenation: {nonHyphenatedTimeMs} ms");
    }

    private static Document CreateSampleDocument(bool autoHyphenation)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Narrow page width forces line wrapping, making hyphenation visible.
        doc.FirstSection.PageSetup.PageWidth = 300; // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Sample text containing words that can be hyphenated.
        builder.Font.Size = 12;
        builder.Writeln(
            "extraordinarycharacteristically internationalization communication " +
            "extraordinarycharacteristically internationalization communication " +
            "extraordinarycharacteristically internationalization communication.");

        // Configure hyphenation options.
        doc.HyphenationOptions.AutoHyphenation = autoHyphenation;
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch (in 1/20 points)
        doc.HyphenationOptions.HyphenateCaps = true;

        // Set the document locale to match the registered dictionary.
        builder.Font.LocaleId = new CultureInfo("en-US").LCID;

        return doc;
    }
}
