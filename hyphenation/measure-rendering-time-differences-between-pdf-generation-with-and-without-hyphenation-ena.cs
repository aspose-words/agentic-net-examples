using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a minimal hyphenation dictionary for English (US).
        const string dictFileName = "hyph_en_US.dic";
        File.WriteAllText(dictFileName,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Register the dictionary for the "en-US" locale.
        Hyphenation.RegisterDictionary("en-US", dictFileName);

        // Build a document with text long enough to trigger hyphenation.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow page width forces line wrapping.
        doc.FirstSection.PageSetup.PageWidth = 200;
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Set the paragraph locale to match the dictionary language.
        builder.Font.LocaleId = new CultureInfo("en-US").LCID;

        // Sample text containing words present in the dictionary, repeated to ensure wrapping.
        string sample = "extraordinarycharacteristically internationalization communication ";
        for (int i = 0; i < 20; i++)
        {
            builder.Writeln(sample);
        }

        // Measure PDF generation with automatic hyphenation enabled.
        doc.HyphenationOptions.AutoHyphenation = true;
        Stopwatch sw = new Stopwatch();
        sw.Start();
        doc.Save("hyphenated.pdf");
        sw.Stop();
        long hyphenatedTime = sw.ElapsedMilliseconds;
        sw.Reset();

        // Measure PDF generation with hyphenation disabled.
        doc.HyphenationOptions.AutoHyphenation = false;
        // Force layout recomputation after changing hyphenation settings.
        doc.UpdatePageLayout();
        sw.Start();
        doc.Save("nonhyphenated.pdf");
        sw.Stop();
        long nonHyphenatedTime = sw.ElapsedMilliseconds;

        // Validate that the output files were created.
        if (!File.Exists("hyphenated.pdf"))
            throw new InvalidOperationException("Hyphenated PDF was not created.");
        if (!File.Exists("nonhyphenated.pdf"))
            throw new InvalidOperationException("Non‑hyphenated PDF was not created.");

        // Output the measured times.
        Console.WriteLine($"Hyphenated PDF generation time: {hyphenatedTime} ms");
        Console.WriteLine($"Non‑hyphenated PDF generation time: {nonHyphenatedTime} ms");
    }
}
