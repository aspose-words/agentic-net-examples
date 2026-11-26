using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class HyphenationPdfComparison
{
    public static void Main()
    {
        // File paths
        const string nonHyphenatedPdf = "nonhyphenated.pdf";
        const string hyphenatedPdf = "hyphenated.pdf";
        const string dictionaryPath = "hyph_en_US.dic";

        // Ensure any previous files are removed
        if (File.Exists(nonHyphenatedPdf)) File.Delete(nonHyphenatedPdf);
        if (File.Exists(hyphenatedPdf)) File.Delete(hyphenatedPdf);
        if (File.Exists(dictionaryPath)) File.Delete(dictionaryPath);

        // Create and save a document without hyphenation
        Document nonHyphenatedDoc = CreateSampleDocument();
        nonHyphenatedDoc.Save(nonHyphenatedPdf, SaveFormat.Pdf);
        ValidateFileExists(nonHyphenatedPdf);

        // Create a minimal hyphenation dictionary
        const string dictionaryContent =
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n";
        File.WriteAllText(dictionaryPath, dictionaryContent);

        // Register the dictionary for en-US locale
        Hyphenation.RegisterDictionary("en-US", dictionaryPath);

        // Create and save a document with hyphenation enabled
        Document hyphenatedDoc = CreateSampleDocument();
        hyphenatedDoc.Save(hyphenatedPdf, SaveFormat.Pdf);
        ValidateFileExists(hyphenatedPdf);

        // Compare file sizes
        long nonSize = new FileInfo(nonHyphenatedPdf).Length;
        long hyphenSize = new FileInfo(hyphenatedPdf).Length;

        Console.WriteLine($"Non‑hyphenated PDF size: {nonSize} bytes");
        Console.WriteLine($"Hyphenated PDF size: {hyphenSize} bytes");

        if (hyphenSize < nonSize)
        {
            Console.WriteLine("Hyphenated PDF is smaller than the non‑hyphenated version.");
        }
        else if (hyphenSize > nonSize)
        {
            Console.WriteLine("Hyphenated PDF is larger than the non‑hyphenated version.");
        }
        else
        {
            Console.WriteLine("Both PDFs have the same size.");
        }
    }

    // Creates a sample document with narrow page width to force line wrapping.
    private static Document CreateSampleDocument()
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow page width to increase the chance of hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 200; // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Sample text containing words that can be hyphenated.
        builder.Writeln("extraordinarycharacteristically internationalization communication");

        return doc;
    }

    // Throws if the expected file does not exist.
    private static void ValidateFileExists(string path)
    {
        if (!File.Exists(path))
        {
            throw new InvalidOperationException($"Expected output file '{path}' was not created.");
        }
    }
}
