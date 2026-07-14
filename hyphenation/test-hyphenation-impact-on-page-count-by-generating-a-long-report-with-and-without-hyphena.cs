using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class HyphenationPageCountDemo
{
    public static void Main()
    {
        // Prepare a minimal hyphenation dictionary for English (US).
        const string dictFileName = "hyph_en_US.dic";
        const string dictContent =
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n";

        File.WriteAllText(dictFileName, dictContent);

        // Register the dictionary for the "en-US" locale.
        Hyphenation.RegisterDictionary("en-US", dictFileName);

        // Generate a long text that will be used for both documents.
        const string longParagraph =
            "extraordinarycharacteristically internationalization communication " +
            "extraordinarycharacteristically internationalization communication " +
            "extraordinarycharacteristically internationalization communication.";

        // Create a document without hyphenation.
        var docNoHyphen = new Document();
        var builder = new DocumentBuilder(docNoHyphen);
        ConfigurePageSetup(docNoHyphen);
        builder.Font.Size = 12;
        builder.Font.LocaleId = new CultureInfo("en-US").LCID;

        // Write enough paragraphs to span multiple pages.
        for (int i = 0; i < 50; i++)
            builder.Writeln(longParagraph);

        const string noHyphenFile = "NoHyphenation.pdf";
        docNoHyphen.Save(noHyphenFile, SaveFormat.Pdf);
        if (!File.Exists(noHyphenFile))
            throw new InvalidOperationException("Failed to create the document without hyphenation.");

        int pagesWithoutHyphen = docNoHyphen.PageCount;

        // Create a document with automatic hyphenation enabled.
        var docWithHyphen = new Document();
        builder = new DocumentBuilder(docWithHyphen);
        ConfigurePageSetup(docWithHyphen);
        builder.Font.Size = 12;
        builder.Font.LocaleId = new CultureInfo("en-US").LCID;

        // Enable hyphenation.
        docWithHyphen.HyphenationOptions.AutoHyphenation = true;
        docWithHyphen.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        docWithHyphen.HyphenationOptions.HyphenationZone = 720; // 0.5 inch

        for (int i = 0; i < 50; i++)
            builder.Writeln(longParagraph);

        const string withHyphenFile = "WithHyphenation.pdf";
        docWithHyphen.Save(withHyphenFile, SaveFormat.Pdf);
        if (!File.Exists(withHyphenFile))
            throw new InvalidOperationException("Failed to create the document with hyphenation.");

        int pagesWithHyphen = docWithHyphen.PageCount;

        // Output the comparison.
        Console.WriteLine($"Pages without hyphenation: {pagesWithoutHyphen}");
        Console.WriteLine($"Pages with hyphenation   : {pagesWithHyphen}");

        // Clean up the temporary dictionary file.
        File.Delete(dictFileName);
    }

    private static void ConfigurePageSetup(Document doc)
    {
        // Use a narrow page width to force line wrapping.
        doc.FirstSection.PageSetup.PageWidth = 300; // points (~4.17 inches)
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;
    }
}
