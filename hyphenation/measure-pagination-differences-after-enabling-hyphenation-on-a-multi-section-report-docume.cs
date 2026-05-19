using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class HyphenationPaginationDemo
{
    public static void Main()
    {
        // Prepare a deterministic output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a minimal hyphenation dictionary for English (US).
        string dictPath = Path.Combine(outputDir, "hyph_en_US.dic");
        File.WriteAllText(dictPath,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n" +
            "hyphenation=hy-phen-a-tion\n");

        // Register the dictionary and verify registration.
        Hyphenation.RegisterDictionary("en-US", dictPath);
        if (!Hyphenation.IsDictionaryRegistered("en-US"))
            throw new InvalidOperationException("Failed to register hyphenation dictionary.");

        // Build a multi‑section document with long text that can be hyphenated.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow the page width to force line wrapping.
        doc.FirstSection.PageSetup.PageWidth = 300; // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Use English (US) locale for hyphenation.
        builder.Font.LocaleId = new CultureInfo("en-US").LCID;
        builder.Font.Size = 12;

        string longText = "extraordinarycharacteristically internationalization communication hyphenation demonstration " +
                          "extraordinarycharacteristically internationalization communication hyphenation demonstration " +
                          "extraordinarycharacteristically internationalization communication hyphenation demonstration.";

        // Section 1
        builder.Writeln(longText);
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Section 2
        builder.Writeln(longText);
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Section 3
        builder.Writeln(longText);

        // Save the document without hyphenation.
        string noHyphenPath = Path.Combine(outputDir, "NoHyphenation.pdf");
        doc.Save(noHyphenPath, SaveFormat.Pdf);
        if (!File.Exists(noHyphenPath))
            throw new InvalidOperationException("Failed to create PDF without hyphenation.");

        int pagesWithoutHyphen = doc.PageCount;

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;
        // Set a valid hyphenation zone (default is 360 = 0.25 inch). Using the default value avoids the exception.
        doc.HyphenationOptions.HyphenationZone = 360;
        doc.UpdatePageLayout();

        // Save the document with hyphenation.
        string hyphenPath = Path.Combine(outputDir, "WithHyphenation.pdf");
        doc.Save(hyphenPath, SaveFormat.Pdf);
        if (!File.Exists(hyphenPath))
            throw new InvalidOperationException("Failed to create PDF with hyphenation.");

        int pagesWithHyphen = doc.PageCount;

        // Output the pagination comparison.
        Console.WriteLine($"Pages without hyphenation: {pagesWithoutHyphen}");
        Console.WriteLine($"Pages with hyphenation   : {pagesWithHyphen}");
    }
}
