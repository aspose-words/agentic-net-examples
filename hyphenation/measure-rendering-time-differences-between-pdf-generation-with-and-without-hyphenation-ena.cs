using System;
using System.Diagnostics;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a minimal hyphenation dictionary for English (US).
        const string dictFile = "hyph_en_US.dic";
        File.WriteAllText(dictFile,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Register the dictionary so that hyphenation can be applied.
        Hyphenation.RegisterDictionary("en-US", dictFile);

        // Measure PDF generation without hyphenation.
        Document docWithoutHyphen = CreateSampleDocument();
        docWithoutHyphen.HyphenationOptions.AutoHyphenation = false;
        const string pdfWithout = "no_hyphenation.pdf";
        var sw = Stopwatch.StartNew();
        docWithoutHyphen.Save(pdfWithout);
        sw.Stop();
        long timeWithout = sw.ElapsedMilliseconds;
        if (!File.Exists(pdfWithout))
            throw new InvalidOperationException("PDF without hyphenation was not created.");

        // Measure PDF generation with hyphenation.
        Document docWithHyphen = CreateSampleDocument();
        docWithHyphen.HyphenationOptions.AutoHyphenation = true;
        const string pdfWith = "hyphenation.pdf";
        sw.Restart();
        docWithHyphen.Save(pdfWith);
        sw.Stop();
        long timeWith = sw.ElapsedMilliseconds;
        if (!File.Exists(pdfWith))
            throw new InvalidOperationException("PDF with hyphenation was not created.");

        // Output the timing results.
        Console.WriteLine($"PDF generation time without hyphenation: {timeWithout} ms");
        Console.WriteLine($"PDF generation time with hyphenation:    {timeWith} ms");
    }

    // Creates a sample document containing long text that will wrap and potentially hyphenate.
    private static Document CreateSampleDocument()
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Narrow page width forces line wrapping.
        doc.FirstSection.PageSetup.PageWidth = 300;
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Sample text with long words that can be hyphenated.
        string sampleText = "extraordinarycharacteristically internationalization communication " +
                            "extraordinarycharacteristically internationalization communication " +
                            "extraordinarycharacteristically internationalization communication.";

        // Write the text multiple times to create several pages.
        for (int i = 0; i < 5; i++)
        {
            builder.Writeln(sampleText);
        }

        return doc;
    }
}
