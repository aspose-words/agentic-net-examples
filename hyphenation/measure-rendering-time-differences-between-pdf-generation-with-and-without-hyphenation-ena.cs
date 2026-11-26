using System;
using System.Diagnostics;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class HyphenationTimingExample
{
    private const string DictionaryFileName = "hyph_en_US.dic";
    private const string HyphenatedPdf = "hyphenated.pdf";
    private const string NonHyphenatedPdf = "non_hyphenated.pdf";

    public static void Main()
    {
        // Ensure Aspose.Words license is not required for this example.
        CreateDictionaryFile();

        // Create and render document with hyphenation enabled.
        RegisterDictionary();
        Document hyphenatedDoc = CreateSampleDocument();
        double hyphenatedTimeMs = RenderPdf(hyphenatedDoc, HyphenatedPdf);
        ValidateFileExists(HyphenatedPdf);

        // Create and render document without hyphenation.
        // Do not register dictionary for this run.
        Document nonHyphenatedDoc = CreateSampleDocument();
        double nonHyphenatedTimeMs = RenderPdf(nonHyphenatedDoc, NonHyphenatedPdf);
        ValidateFileExists(NonHyphenatedPdf);

        // Output timing results.
        Console.WriteLine($"PDF generation with hyphenation: {hyphenatedTimeMs:F2} ms");
        Console.WriteLine($"PDF generation without hyphenation: {nonHyphenatedTimeMs:F2} ms");
    }

    private static void CreateDictionaryFile()
    {
        string[] lines =
        {
            "UTF-8",
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly",
            "internationalization=in-ter-na-tion-al-i-za-tion",
            "communication=com-mu-ni-ca-tion"
        };
        File.WriteAllLines(DictionaryFileName, lines);
        if (!File.Exists(DictionaryFileName))
            throw new InvalidOperationException("Failed to create hyphenation dictionary file.");
    }

    private static void RegisterDictionary()
    {
        // Register the dictionary for en-US language.
        Hyphenation.RegisterDictionary("en-US", DictionaryFileName);
    }

    private static Document CreateSampleDocument()
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow page width to force line wrapping.
        Section section = doc.FirstSection;
        section.PageSetup.PageWidth = 200; // points
        section.PageSetup.LeftMargin = 20;
        section.PageSetup.RightMargin = 20;

        // Add sample text containing long words.
        builder.Font.Size = 12;
        builder.Writeln("extraordinarycharacteristically internationalization communication");
        builder.Writeln("extraordinarycharacteristically internationalization communication");
        builder.Writeln("extraordinarycharacteristically internationalization communication");

        return doc;
    }

    private static double RenderPdf(Document doc, string outputPath)
    {
        Stopwatch sw = Stopwatch.StartNew();
        doc.Save(outputPath, SaveFormat.Pdf);
        sw.Stop();
        return sw.Elapsed.TotalMilliseconds;
    }

    private static void ValidateFileExists(string path)
    {
        if (!File.Exists(path))
            throw new InvalidOperationException($"Expected output file '{path}' was not created.");
    }
}
