using System;
using System.Globalization;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the text locale to German (Switzerland) so that Word uses German hyphenation rules.
        builder.Font.LocaleId = new CultureInfo("de-CH").LCID;

        // Write a long German compound word that will need hyphenation.
        builder.Writeln("Donaudampfschifffahrtsgesellschaftskapitän");

        // Narrow the page width to force line wrapping and hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 200;
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Enable automatic hyphenation for the document.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Create a minimal German hyphenation dictionary file.
        const string dictPath = "hyph_de_CH.dic";
        File.WriteAllText(dictPath,
            "UTF-8\n" +
            "Donaudampfschifffahrtsgesellschaftskapitän=Do-nau-dampf-schiff-fahrts-ge-sell-schafts-ka-pit-än");

        // Register the dictionary for the \"de-CH\" language code.
        Hyphenation.RegisterDictionary("de-CH", dictPath);

        // Save the document as PDF to observe hyphenation.
        const string outPath = "HyphenatedGerman.pdf";
        doc.Save(outPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(outPath))
            throw new InvalidOperationException("Expected PDF output was not created.");
    }
}
