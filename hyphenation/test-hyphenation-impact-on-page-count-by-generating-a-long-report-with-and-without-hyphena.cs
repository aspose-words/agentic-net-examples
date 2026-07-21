using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;
using Aspose.Words.Saving;

public class HyphenationPageCountDemo
{
    public static void Main()
    {
        // Create a minimal hyphenation dictionary for English (US).
        const string dictFileName = "hyph_en_US.dic";
        const string dictContent =
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n";
        File.WriteAllText(dictFileName, dictContent);

        // Register the dictionary so that Aspose.Words can hyphenate English text.
        Hyphenation.RegisterDictionary("en-US", dictFileName);

        // Build a long paragraph that will wrap many times on a narrow page.
        string repeatedSentence = "extraordinarycharacteristically internationalization communication";
        string longText = string.Join(" ", System.Linq.Enumerable.Repeat(repeatedSentence, 120));

        // -----------------------------------------------------------------
        // Document with automatic hyphenation enabled.
        // -----------------------------------------------------------------
        Document hyphenatedDoc = new Document();
        DocumentBuilder hyBuilder = new DocumentBuilder(hyphenatedDoc);

        // Narrow page to force many line breaks.
        hyphenatedDoc.FirstSection.PageSetup.PageWidth = 300; // points (~4.2 inches)
        hyphenatedDoc.FirstSection.PageSetup.LeftMargin = 20;
        hyphenatedDoc.FirstSection.PageSetup.RightMargin = 20;

        // Set font and locale for hyphenation.
        hyBuilder.Font.Size = 12;
        hyBuilder.Font.LocaleId = new CultureInfo("en-US").LCID;

        // Write the long text.
        hyBuilder.Writeln(longText);

        // Enable automatic hyphenation.
        hyphenatedDoc.HyphenationOptions.AutoHyphenation = true;
        // Use the default hyphenation zone (no need to set it to 0, which is invalid).
        hyphenatedDoc.HyphenationOptions.HyphenateCaps = true;
        hyphenatedDoc.HyphenationOptions.ConsecutiveHyphenLimit = 0;

        // Save as PDF and verify creation.
        const string hyphenatedPdf = "Hyphenated.pdf";
        hyphenatedDoc.Save(hyphenatedPdf, SaveFormat.Pdf);
        if (!File.Exists(hyphenatedPdf))
            throw new InvalidOperationException("Hyphenated PDF was not created.");

        int hyphenatedPages = hyphenatedDoc.PageCount;

        // -----------------------------------------------------------------
        // Document with hyphenation disabled.
        // -----------------------------------------------------------------
        Document nonHyphenatedDoc = new Document();
        DocumentBuilder nonBuilder = new DocumentBuilder(nonHyphenatedDoc);

        nonHyphenatedDoc.FirstSection.PageSetup.PageWidth = 300;
        nonHyphenatedDoc.FirstSection.PageSetup.LeftMargin = 20;
        nonHyphenatedDoc.FirstSection.PageSetup.RightMargin = 20;

        nonBuilder.Font.Size = 12;
        nonBuilder.Font.LocaleId = new CultureInfo("en-US").LCID;
        nonBuilder.Writeln(longText);

        // Ensure hyphenation is turned off (default is false, but set explicitly).
        nonHyphenatedDoc.HyphenationOptions.AutoHyphenation = false;

        const string nonHyphenatedPdf = "NonHyphenated.pdf";
        nonHyphenatedDoc.Save(nonHyphenatedPdf, SaveFormat.Pdf);
        if (!File.Exists(nonHyphenatedPdf))
            throw new InvalidOperationException("Non‑hyphenated PDF was not created.");

        int nonHyphenatedPages = nonHyphenatedDoc.PageCount;

        // Output the comparison result.
        Console.WriteLine($"Pages with hyphenation    : {hyphenatedPages}");
        Console.WriteLine($"Pages without hyphenation : {nonHyphenatedPages}");
        Console.WriteLine(hyphenatedPages < nonHyphenatedPages
            ? "Hyphenation reduced the page count."
            : hyphenatedPages > nonHyphenatedPages
                ? "Hyphenation increased the page count."
                : "Hyphenation did not change the page count.");
    }
}
