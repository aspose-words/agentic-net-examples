using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define file names in the current working directory.
        string dictionaryPath = Path.Combine(Directory.GetCurrentDirectory(), "hyph_en_US.dic");
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Hyphenated.docx");

        // Create a minimal hyphenation dictionary for English (US).
        // First line is the encoding, subsequent lines are word=hyphenation-points.
        File.WriteAllText(dictionaryPath,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Register the dictionary so Aspose.Words can hyphenate words of this language.
        Hyphenation.RegisterDictionary("en-US", dictionaryPath);
        if (!Hyphenation.IsDictionaryRegistered("en-US"))
            throw new InvalidOperationException("Failed to register the hyphenation dictionary.");

        // Create a new blank document and a builder to add content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow the page width to force line wrapping and hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 200; // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Write a paragraph containing a long word that can be hyphenated.
        builder.Font.Size = 12;
        builder.Writeln("The word extraordinarycharacteristically demonstrates hyphenation when the line is too short.");

        // Enable automatic hyphenation and configure optional settings.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch (720 / 20 = 36 points)
        doc.HyphenationOptions.HyphenateCaps = true;

        // Save the document to DOCX format.
        doc.Save(outputPath);
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output DOCX file was not created.");

        // Load the saved document to verify that hyphenation settings persisted.
        Document loadedDoc = new Document(outputPath);
        if (!loadedDoc.HyphenationOptions.AutoHyphenation)
            throw new InvalidOperationException("Auto hyphenation option was not retained after saving.");

        // Verify that the dictionary is still registered (required for hyphenation on load).
        if (!Hyphenation.IsDictionaryRegistered("en-US"))
            throw new InvalidOperationException("Hyphenation dictionary is not registered after loading the document.");

        // Simple verification that the long word exists in the document text.
        string text = loadedDoc.GetText();
        if (!text.Contains("extraordinarycharacteristically"))
            throw new InvalidOperationException("The expected text was not found in the loaded document.");

        // Optional clean‑up (commented out to allow inspection of the generated files).
        // File.Delete(dictionaryPath);
        // File.Delete(outputPath);
    }
}
