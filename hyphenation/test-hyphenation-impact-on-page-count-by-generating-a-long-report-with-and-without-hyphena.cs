using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a minimal hyphenation dictionary for English (en-US).
        const string dictionaryFileName = "hyph_en_US.dic";
        File.WriteAllText(dictionaryFileName,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Create a long report without hyphenation.
        Document docWithoutHyphenation = CreateReport();
        const string noHyphenPdf = "Report_NoHyphenation.pdf";
        docWithoutHyphenation.Save(noHyphenPdf, SaveFormat.Pdf);
        ValidateFileExists(noHyphenPdf);
        docWithoutHyphenation.UpdatePageLayout();
        int pagesWithout = docWithoutHyphenation.PageCount;

        // Register the dictionary and create the same report with hyphenation enabled.
        Hyphenation.RegisterDictionary("en-US", dictionaryFileName);
        Document docWithHyphenation = CreateReport();
        const string withHyphenPdf = "Report_WithHyphenation.pdf";
        docWithHyphenation.Save(withHyphenPdf, SaveFormat.Pdf);
        ValidateFileExists(withHyphenPdf);
        docWithHyphenation.UpdatePageLayout();
        int pagesWith = docWithHyphenation.PageCount;

        // Output the page counts for comparison.
        Console.WriteLine($"Pages without hyphenation: {pagesWithout}");
        Console.WriteLine($"Pages with hyphenation: {pagesWith}");
    }

    // Creates a document containing repeated sample text.
    private static Document CreateReport()
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow page width forces many line breaks.
        Section section = doc.FirstSection;
        section.PageSetup.PageWidth = 200; // points
        section.PageSetup.LeftMargin = 20;
        section.PageSetup.RightMargin = 20;

        // Sample text that matches entries in the dictionary.
        string sampleSentence = "extraordinarycharacteristically internationalization communication ";

        // Repeat the sentence enough times to span several pages.
        for (int i = 0; i < 200; i++)
        {
            builder.Writeln(sampleSentence);
        }

        return doc;
    }

    // Throws if the expected file was not created.
    private static void ValidateFileExists(string path)
    {
        if (!File.Exists(path))
            throw new InvalidOperationException($"Expected output file '{path}' was not created.");
    }
}
