using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Path for the German hyphenation dictionary.
        const string dictionaryPath = "hyph_de_DE.dic";

        // Create a minimal dictionary that contains a compound‑word pattern.
        // The first line must specify the encoding, followed by entries in the form word=pattern.
        File.WriteAllText(dictionaryPath,
            "UTF-8\n" +
            "Donaudampfschifffahrtsgesellschaft=Do-nau-dampf-schiff-fahrts-ge-sell-schaft\n");

        // Register the dictionary for the German locale.
        Hyphenation.RegisterDictionary("de-DE", dictionaryPath);

        // Verify that the dictionary was registered.
        if (!Hyphenation.IsDictionaryRegistered("de-DE"))
            throw new InvalidOperationException("German hyphenation dictionary was not registered.");

        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the locale of the text to German.
        builder.Font.LocaleId = new CultureInfo("de-DE").LCID;
        builder.Font.Size = 12;

        // Add a paragraph containing a long German compound word.
        builder.Writeln(
            "Donaudampfschifffahrtsgesellschaft ist ein sehr langes deutsches Wort, das hypheniert werden soll.");

        // Narrow the page width so that hyphenation can be observed.
        doc.FirstSection.PageSetup.PageWidth = 200;
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Save the document as PDF.
        const string outputPath = "HyphenatedGerman.pdf";
        doc.Save(outputPath);

        // Validate that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The PDF file was not created.");
    }
}
