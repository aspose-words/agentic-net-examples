using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Prepare a minimal hyphenation dictionary for English (US).
        const string dictFileName = "hyph_en_US.dic";
        File.WriteAllText(dictFileName,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Register the dictionary so that hyphenation can be applied.
        Hyphenation.RegisterDictionary("en-US", dictFileName);

        // Create a sample document with long words that can be hyphenated.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Size = 12;
        builder.Writeln(
            "extraordinarycharacteristically internationalization communication " +
            "extraordinarycharacteristically internationalization communication " +
            "extraordinarycharacteristically internationalization communication.");

        // Narrow the page width to force line wrapping and thus hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 300;
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Measure PDF generation with automatic hyphenation enabled.
        doc.HyphenationOptions.AutoHyphenation = true;
        string hyphenatedPdf = "Hyphenated.pdf";
        Stopwatch sw = Stopwatch.StartNew();
        doc.Save(hyphenatedPdf, SaveFormat.Pdf);
        sw.Stop();
        long timeWithHyphenation = sw.ElapsedMilliseconds;

        // Verify the PDF was created.
        if (!File.Exists(hyphenatedPdf))
            throw new InvalidOperationException("Hyphenated PDF was not created.");

        // Measure PDF generation with automatic hyphenation disabled.
        doc.HyphenationOptions.AutoHyphenation = false;
        string nonHyphenatedPdf = "NonHyphenated.pdf";
        sw.Restart();
        doc.Save(nonHyphenatedPdf, SaveFormat.Pdf);
        sw.Stop();
        long timeWithoutHyphenation = sw.ElapsedMilliseconds;

        // Verify the second PDF was created.
        if (!File.Exists(nonHyphenatedPdf))
            throw new InvalidOperationException("Non‑hyphenated PDF was not created.");

        // Output the timing results.
        Console.WriteLine($"PDF generation time with hyphenation: {timeWithHyphenation} ms");
        Console.WriteLine($"PDF generation time without hyphenation: {timeWithoutHyphenation} ms");
    }
}
