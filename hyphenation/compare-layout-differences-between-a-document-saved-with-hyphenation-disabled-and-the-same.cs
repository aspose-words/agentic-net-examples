using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

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

        // Register the dictionary so Aspose.Words can hyphenate English text.
        Hyphenation.RegisterDictionary("en-US", dictPath);

        // Build a document with a narrow page width to force line wrapping.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow page setup.
        Section section = doc.FirstSection;
        section.PageSetup.PageWidth = 200;   // points
        section.PageSetup.LeftMargin = 20;
        section.PageSetup.RightMargin = 20;

        // Use English locale for hyphenation.
        builder.Font.LocaleId = new CultureInfo("en-US").LCID;
        builder.Font.Size = 12;

        // Sample text containing long words that can be hyphenated.
        builder.Writeln("extraordinarycharacteristically internationalization communication");

        // Save the document with hyphenation disabled (default).
        const string disabledPath = "HyphenationDisabled.pdf";
        doc.Save(disabledPath, SaveFormat.Pdf);
        if (!File.Exists(disabledPath))
            throw new InvalidOperationException("Failed to create the disabled‑hyphenation PDF.");

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch (in 1/20 pt units)
        doc.HyphenationOptions.HyphenateCaps = true;

        // Rebuild layout to apply hyphenation changes.
        doc.UpdatePageLayout();

        // Save the document with hyphenation enabled.
        const string enabledPath = "HyphenationEnabled.pdf";
        doc.Save(enabledPath, SaveFormat.Pdf);
        if (!File.Exists(enabledPath))
            throw new InvalidOperationException("Failed to create the enabled‑hyphenation PDF.");

        // Load both PDFs to compare their page counts.
        Document disabledDoc = new Document(disabledPath);
        Document enabledDoc = new Document(enabledPath);

        int disabledPageCount = disabledDoc.PageCount;
        int enabledPageCount = enabledDoc.PageCount;

        // Output the comparison result.
        Console.WriteLine($"Page count with hyphenation disabled: {disabledPageCount}");
        Console.WriteLine($"Page count with hyphenation enabled : {enabledPageCount}");
    }
}
