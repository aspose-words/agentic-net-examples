using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class HyphenationPageCountDemo
{
    private const string DictionaryFileName = "hyph_en_US.dic";
    private const string OutputWithoutHyphenation = "report_without_hyphenation.pdf";
    private const string OutputWithHyphenation = "report_with_hyphenation.pdf";

    public static void Main()
    {
        // Ensure the hyphenation dictionary exists.
        CreateDictionaryFile();

        // Generate report without hyphenation.
        int pagesWithout = GenerateReport(enableHyphenation: false, OutputWithoutHyphenation);
        // Generate report with hyphenation.
        int pagesWith = GenerateReport(enableHyphenation: true, OutputWithHyphenation);

        // Validate that the PDF files were created.
        if (!File.Exists(OutputWithoutHyphenation))
            throw new InvalidOperationException($"File '{OutputWithoutHyphenation}' was not created.");
        if (!File.Exists(OutputWithHyphenation))
            throw new InvalidOperationException($"File '{OutputWithHyphenation}' was not created.");

        // Output the page counts.
        Console.WriteLine($"Pages without hyphenation: {pagesWithout}");
        Console.WriteLine($"Pages with hyphenation:    {pagesWith}");

        // Simple verification: hyphenation should not increase page count.
        if (pagesWith > pagesWithout)
            throw new InvalidOperationException("Hyphenation increased the page count, which is unexpected for this test.");
    }

    private static void CreateDictionaryFile()
    {
        // Minimal dictionary content for English (US) hyphenation.
        // The first line must be the encoding identifier.
        string content =
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n";

        File.WriteAllText(DictionaryFileName, content);
        // Register the dictionary for the "en-US" locale.
        Hyphenation.RegisterDictionary("en-US", DictionaryFileName);
    }

    private static int GenerateReport(bool enableHyphenation, string outputPath)
    {
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Set a narrow page width to force many line wraps.
        doc.FirstSection.PageSetup.PageWidth = 300; // points (~4.17 inches)
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Use English (US) locale for hyphenation.
        builder.Font.LocaleId = new CultureInfo("en-US").LCID;
        builder.Font.Size = 12;

        // Add a large amount of text that contains hyphenatable words.
        for (int i = 0; i < 200; i++)
        {
            builder.Writeln(
                "extraordinarycharacteristically internationalization communication " +
                "extraordinarycharacteristically internationalization communication");
        }

        if (enableHyphenation)
        {
            // Enable automatic hyphenation for the document.
            doc.HyphenationOptions.AutoHyphenation = true;
            doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
            doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch
            doc.HyphenationOptions.HyphenateCaps = true;
        }

        // Save the document; this also triggers layout calculation.
        doc.Save(outputPath, SaveFormat.Pdf);

        // Return the calculated page count.
        return doc.PageCount;
    }
}
