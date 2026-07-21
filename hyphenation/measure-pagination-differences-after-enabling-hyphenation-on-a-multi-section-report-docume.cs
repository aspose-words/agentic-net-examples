using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class HyphenationPaginationDemo
{
    public static void Main()
    {
        // Prepare a minimal hyphenation dictionary for English (US).
        const string dictFileName = "hyph_en_US.dic";
        const string dictContent =
            "UTF-8\n" +
            "extraordinarycharacteristically=ex-tra-or-di-na-ry-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n";
        File.WriteAllText(dictFileName, dictContent);

        // Register the dictionary.
        Hyphenation.RegisterDictionary("en-US", dictFileName);

        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow page width to force line wrapping.
        doc.FirstSection.PageSetup.PageWidth = 300; // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Sample long text that can be hyphenated.
        string sampleText = "extraordinarycharacteristically internationalization communication ";

        // Build two sections with the same content.
        for (int sec = 1; sec <= 2; sec++)
        {
            // Set the locale for hyphenation.
            builder.Font.LocaleId = new CultureInfo("en-US").LCID;
            builder.Writeln($"Section {sec}:");
            // Write several paragraphs to increase page count.
            for (int p = 0; p < 5; p++)
            {
                builder.Writeln(sampleText + sampleText + sampleText);
            }

            if (sec < 2)
                builder.InsertBreak(BreakType.SectionBreakNewPage);
        }

        // Measure pagination before enabling hyphenation.
        int pagesWithoutHyphenation = doc.PageCount;
        const string pdfWithoutHyphen = "report_without_hyphenation.pdf";
        doc.Save(pdfWithoutHyphen, SaveFormat.Pdf);
        if (!File.Exists(pdfWithoutHyphen))
            throw new InvalidOperationException("Failed to create PDF without hyphenation.");

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch
        doc.HyphenationOptions.HyphenateCaps = true;

        // Recalculate layout after changing hyphenation settings.
        doc.UpdatePageLayout();

        // Measure pagination after hyphenation.
        int pagesWithHyphenation = doc.PageCount;
        const string pdfWithHyphen = "report_with_hyphenation.pdf";
        doc.Save(pdfWithHyphen, SaveFormat.Pdf);
        if (!File.Exists(pdfWithHyphen))
            throw new InvalidOperationException("Failed to create PDF with hyphenation.");

        // Output the results.
        Console.WriteLine($"Pages without hyphenation: {pagesWithoutHyphenation}");
        Console.WriteLine($"Pages with hyphenation:    {pagesWithHyphenation}");
        Console.WriteLine($"Page count difference:    {pagesWithoutHyphenation - pagesWithHyphenation}");
    }
}
