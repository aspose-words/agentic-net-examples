using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class HyphenationCompoundWordExample
{
    public static void Main()
    {
        // Paths for temporary files.
        const string dictionaryPath = "hyph_de_CH.dic";
        const string outputPath = "GermanHyphenated.pdf";

        // Create a minimal German hyphenation dictionary.
        // The first line must specify the encoding, followed by word=hyphenation patterns.
        // The pattern uses hyphens to indicate allowed break points.
        File.WriteAllText(dictionaryPath,
            "UTF-8\n" +
            "Donaudampfschifffahrtsgesellschaft=Do-nau-dampf-schiff-fahrts-ge-sell-schaft\n" +
            "Bundesverfassungsgericht=Bundes-ver-fas-sungs-ge-richt\n");

        // Register the German dictionary.
        Hyphenation.RegisterDictionary("de-CH", dictionaryPath);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set a narrow page width to force line wrapping.
        doc.FirstSection.PageSetup.PageWidth = 300; // points (~4.2 cm)
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;
        // Hyphenate as close to the margin as possible – use a small positive value.
        doc.HyphenationOptions.HyphenationZone = 1;

        // Set the language of the text to German (Switzerland) – the same code used for the dictionary.
        builder.Font.LocaleId = new CultureInfo("de-CH").LCID;

        // Write a paragraph containing long German compound words that require hyphenation.
        builder.Writeln(
            "Die Donaudampfschifffahrtsgesellschaft ist ein sehr langes Wort, " +
            "das in der deutschen Sprache häufig hypheniert werden muss, " +
            "um den Textfluss zu erhalten. " +
            "Ein weiteres Beispiel ist das Bundesverfassungsgericht, " +
            "das ebenfalls aus vielen Bestandteilen besteht.");

        // Save the document to PDF.
        doc.Save(outputPath, SaveFormat.Pdf);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"The expected output file '{outputPath}' was not created.");

        // Clean up temporary dictionary file.
        File.Delete(dictionaryPath);
    }
}
