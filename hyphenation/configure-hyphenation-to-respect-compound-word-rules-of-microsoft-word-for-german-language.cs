using System;
using System.Globalization;
using System.IO;
using Aspose.Words;

public class HyphenationGermanExample
{
    public static void Main()
    {
        // Create a minimal German (Switzerland) hyphenation dictionary.
        const string dictFileName = "hyph_de_CH.dic";
        const string dictContent = @"UTF-8
Donaudampfschifffahrtsgesellschaftskapitän=Do-nau-dampf-schiff-fahrts-gesell-schafts-ka-pit-än
";
        File.WriteAllText(dictFileName, dictContent);

        // Register the dictionary for the "de-CH" locale.
        Aspose.Words.Hyphenation.RegisterDictionary("de-CH", dictFileName);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow page width forces line wrapping so hyphenation can be observed.
        doc.FirstSection.PageSetup.PageWidth = 200; // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.HyphenateCaps = true;

        // Set the font locale to German (Switzerland) to match the dictionary.
        builder.Font.LocaleId = new CultureInfo("de-CH").LCID;
        builder.Font.Size = 12;

        // Write a sentence containing a long German compound word.
        builder.Writeln(
            "Die Donaudampfschifffahrtsgesellschaftskapitänin steuerte das Schiff durch die engen Kanäle.");

        // Save the document as PDF.
        const string outputPath = "HyphenatedGerman.pdf";
        doc.Save(outputPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The expected PDF file was not created.");

        // Clean up the temporary dictionary file.
        File.Delete(dictFileName);
    }
}
