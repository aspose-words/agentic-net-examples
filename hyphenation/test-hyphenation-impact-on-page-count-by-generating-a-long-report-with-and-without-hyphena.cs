using System;
using System.Globalization;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Settings;

public class HyphenationPageCountDemo
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

        // Register the dictionary so that Aspose.Words can hyphenate English text.
        Hyphenation.RegisterDictionary("en-US", dictFileName);
        if (!Hyphenation.IsDictionaryRegistered("en-US"))
            throw new InvalidOperationException("Failed to register the hyphenation dictionary.");

        // Build a long paragraph containing words that can be hyphenated.
        string sampleSentence = "extraordinarycharacteristically internationalization communication";
        string longText = string.Join(" ", Enumerable.Repeat(sampleSentence, 50));

        // Create two documents: one without hyphenation and one with hyphenation enabled.
        Document docWithoutHyphenation = CreateDocument(longText, autoHyphenation: false);
        Document docWithHyphenation = CreateDocument(longText, autoHyphenation: true);

        // Save both documents as PDF files.
        const string withoutFile = "ReportWithoutHyphenation.pdf";
        const string withFile = "ReportWithHyphenation.pdf";

        docWithoutHyphenation.Save(withoutFile);
        docWithHyphenation.Save(withFile);

        // Verify that the files were created.
        if (!File.Exists(withoutFile))
            throw new FileNotFoundException("The PDF without hyphenation was not created.", withoutFile);
        if (!File.Exists(withFile))
            throw new FileNotFoundException("The PDF with hyphenation was not created.", withFile);

        // Retrieve page counts. Accessing PageCount forces layout calculation.
        int pagesWithout = docWithoutHyphenation.PageCount;
        int pagesWith = docWithHyphenation.PageCount;

        // Output the results.
        Console.WriteLine($"Pages without hyphenation: {pagesWithout}");
        Console.WriteLine($"Pages with hyphenation   : {pagesWith}");

        // Hyphenation should not increase the page count.
        if (pagesWith > pagesWithout)
            throw new InvalidOperationException("Hyphenation increased the page count, which is unexpected.");

        // Clean up temporary files (optional).
        // File.Delete(dictFileName);
    }

    private static Document CreateDocument(string text, bool autoHyphenation)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a narrow page width to force line wrapping.
        doc.FirstSection.PageSetup.PageWidth = 200; // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Set the locale to English (US) so the registered dictionary is used.
        builder.Font.LocaleId = new CultureInfo("en-US").LCID;
        builder.Font.Size = 12;

        // Write the long text.
        builder.Writeln(text);

        // Configure hyphenation options.
        doc.HyphenationOptions.AutoHyphenation = autoHyphenation;
        // Optional: fine‑tune other hyphenation settings if desired.
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenationZone = 720;
        doc.HyphenationOptions.HyphenateCaps = true;

        return doc;
    }
}
