using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a minimal hyphenation dictionary for English (US).
        string dictPath = Path.Combine(outputDir, "hyph_en_US.dic");
        File.WriteAllText(dictPath,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Register the dictionary.
        Hyphenation.RegisterDictionary("en-US", dictPath);

        // Build a multi‑section document with long text to trigger hyphenation.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow page width forces line wrapping.
        doc.FirstSection.PageSetup.PageWidth = 300; // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Set the locale for hyphenation.
        builder.Font.LocaleId = new CultureInfo("en-US").LCID;

        // Add a paragraph with repetitive long words.
        string longText = "extraordinarycharacteristically internationalization communication ";
        for (int i = 0; i < 30; i++) builder.Writeln(longText);

        // Insert a new section.
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        for (int i = 0; i < 30; i++) builder.Writeln(longText);

        // Ensure layout is calculated before measuring.
        doc.UpdatePageLayout();
        int pagesBefore = doc.PageCount;

        // Save the document without hyphenation.
        string beforePath = Path.Combine(outputDir, "BeforeHyphenation.pdf");
        doc.Save(beforePath, SaveFormat.Pdf);

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch
        doc.HyphenationOptions.HyphenateCaps = true;

        // Re‑layout the document after changing hyphenation settings.
        doc.UpdatePageLayout();
        int pagesAfter = doc.PageCount;

        // Save the hyphenated version.
        string afterPath = Path.Combine(outputDir, "AfterHyphenation.pdf");
        doc.Save(afterPath, SaveFormat.Pdf);

        // Validate that the output files were created.
        if (!File.Exists(beforePath))
            throw new InvalidOperationException("The PDF before hyphenation was not created.");
        if (!File.Exists(afterPath))
            throw new InvalidOperationException("The PDF after hyphenation was not created.");

        // Report the pagination difference.
        Console.WriteLine($"Pages before hyphenation: {pagesBefore}");
        Console.WriteLine($"Pages after hyphenation: {pagesAfter}");
        Console.WriteLine($"Difference: {pagesAfter - pagesBefore}");
    }
}
