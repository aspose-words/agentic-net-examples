using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Paths for the temporary dictionary file and the output document.
        const string dictPath = "hyph_en_US.dic";
        const string docPath = "Hyphenated.docx";

        // Create a minimal hyphenation dictionary file.
        // The format is: first line = encoding, subsequent lines = word=hyphenated‑form.
        string dictContent =
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "communication=com-mu-ni-ca-tion\n";
        File.WriteAllText(dictPath, dictContent);

        // Register the dictionary for the English (US) locale.
        Aspose.Words.Hyphenation.RegisterDictionary("en-US", dictPath);
        if (!Aspose.Words.Hyphenation.IsDictionaryRegistered("en-US"))
            throw new InvalidOperationException("Failed to register hyphenation dictionary.");

        // Create a new blank document and a builder to add content.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Narrow the page width to force line wrapping, making hyphenation visible.
        doc.FirstSection.PageSetup.PageWidth = 200;   // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Enable automatic hyphenation and configure its options.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch (720 / 1440 points per inch)
        doc.HyphenationOptions.HyphenateCaps = true;

        // Add a paragraph containing words that exist in the dictionary.
        builder.Font.Size = 24;
        builder.Writeln("extraordinarycharacteristically communication");

        // Save the document as DOCX.
        doc.Save(docPath, SaveFormat.Docx);
        if (!File.Exists(docPath))
            throw new InvalidOperationException("The DOCX file was not created.");

        // Reload the document to verify that hyphenation settings persisted.
        var loaded = new Document(docPath);
        if (!loaded.HyphenationOptions.AutoHyphenation)
            throw new InvalidOperationException("Auto hyphenation was not retained after reload.");

        // Verify that the dictionary registration is still active.
        if (!Aspose.Words.Hyphenation.IsDictionaryRegistered("en-US"))
            throw new InvalidOperationException("Hyphenation dictionary is no longer registered after reload.");

        // Clean up temporary files (optional).
        // File.Delete(dictPath);
        // File.Delete(docPath);
    }
}
