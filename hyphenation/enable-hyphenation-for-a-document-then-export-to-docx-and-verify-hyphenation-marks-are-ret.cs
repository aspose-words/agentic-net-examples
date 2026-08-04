using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Paths for temporary files
        const string dictionaryPath = "hyph_en_US.dic";
        const string outputPath = "Hyphenated.docx";

        // Clean up any previous files
        if (File.Exists(dictionaryPath))
            File.Delete(dictionaryPath);
        if (File.Exists(outputPath))
            File.Delete(outputPath);

        // Create a minimal hyphenation dictionary for English (US)
        // First line is the encoding, subsequent lines are word=hyphenation-points
        string dictionaryContent =
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n";

        File.WriteAllText(dictionaryPath, dictionaryContent);

        // Register the dictionary with Aspose.Words
        Aspose.Words.Hyphenation.RegisterDictionary("en-US", dictionaryPath);
        if (!Aspose.Words.Hyphenation.IsDictionaryRegistered("en-US"))
            throw new InvalidOperationException("Failed to register hyphenation dictionary.");

        // Create a new document and add long words that can be hyphenated
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Size = 24;
        builder.Writeln("extraordinarycharacteristically internationalization communication");

        // Narrow the page width to force line wrapping and hyphenation
        doc.FirstSection.PageSetup.PageWidth = 200;
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Enable automatic hyphenation
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.HyphenateCaps = true;
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenationZone = 720;

        // Save the document to DOCX
        doc.Save(outputPath);
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output DOCX file was not created.");

        // Reload the saved document to verify hyphenation settings persisted
        Document loaded = new Document(outputPath);
        if (!loaded.HyphenationOptions.AutoHyphenation)
            throw new InvalidOperationException("Hyphenation option was not retained after saving.");

        // Verify the dictionary is still registered (required for hyphenation on load)
        if (!Aspose.Words.Hyphenation.IsDictionaryRegistered("en-US"))
            throw new InvalidOperationException("Hyphenation dictionary is not registered after loading the document.");

        // Optional clean‑up (commented out to allow inspection of the files after run)
        // File.Delete(dictionaryPath);
        // File.Delete(outputPath);
    }
}
