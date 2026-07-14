using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Create a minimal hyphenation dictionary for English (US).
        const string dictionaryFile = "hyph_en_US.dic";
        File.WriteAllText(dictionaryFile,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Register the dictionary so Aspose.Words can hyphenate the words.
        Aspose.Words.Hyphenation.RegisterDictionary("en-US", dictionaryFile);

        // Prepare a document with long words that require hyphenation.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow page width forces line wrapping.
        doc.FirstSection.PageSetup.PageWidth = 200;
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        builder.Font.Size = 12;
        builder.Writeln(
            "extraordinarycharacteristically internationalization communication " +
            "extraordinarycharacteristically internationalization communication");

        // Save the document with hyphenation disabled (default).
        doc.HyphenationOptions.AutoHyphenation = false;
        doc.UpdatePageLayout();
        int pagesWithoutHyphenation = doc.PageCount;
        const string disabledFile = "HyphenationDisabled.pdf";
        doc.Save(disabledFile);
        if (!File.Exists(disabledFile))
            throw new InvalidOperationException("HyphenationDisabled.pdf was not created.");

        // Save the same document with hyphenation enabled.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.UpdatePageLayout();
        int pagesWithHyphenation = doc.PageCount;
        const string enabledFile = "HyphenationEnabled.pdf";
        doc.Save(enabledFile);
        if (!File.Exists(enabledFile))
            throw new InvalidOperationException("HyphenationEnabled.pdf was not created.");

        // Output the layout comparison.
        Console.WriteLine($"Pages without hyphenation: {pagesWithoutHyphenation}");
        Console.WriteLine($"Pages with hyphenation:    {pagesWithHyphenation}");
        Console.WriteLine(pagesWithHyphenation < pagesWithoutHyphenation
            ? "Hyphenation reduced the page count."
            : "Hyphenation did not reduce the page count.");
    }
}
