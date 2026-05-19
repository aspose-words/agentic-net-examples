using System;
using System.Globalization;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the locale to German (Switzerland) so that German hyphenation rules are applied.
        builder.Font.LocaleId = new CultureInfo("de-CH").LCID;

        // Add long German compound words that will need hyphenation when the line wraps.
        builder.Writeln("Donaudampfschifffahrtsgesellschaftskapitänsmütze");
        builder.Writeln("Rindfleischetikettierungsüberwachungsaufgabenübertragungsgesetz");

        // Narrow the page width to force line wrapping.
        doc.FirstSection.PageSetup.PageWidth = 300; // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Enable automatic hyphenation and configure its options.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch (in 1/20 point units)
        doc.HyphenationOptions.HyphenateCaps = true;

        // Create a minimal German hyphenation dictionary file in the current folder.
        string dictFileName = "hyph_de_CH.dic";
        string dictContent =
            "UTF-8\n" +
            "Donaudampfschifffahrtsgesellschaft=Do-nau-dampf-schiff-fahrts-gesell-schaft\n" +
            "Rindfleischetikettierungsüberwachungsaufgabenübertragungsgesetz=Rind-fleisch-etikett-ier-ungs-über-wach-ungs-auf-ga-ben-über-tra-gungs-ge-setz\n";

        File.WriteAllText(dictFileName, dictContent);

        // Register the dictionary for the "de-CH" language.
        Aspose.Words.Hyphenation.RegisterDictionary("de-CH", dictFileName);

        // Save the document as PDF to visualize hyphenation.
        const string outputFile = "HyphenatedGerman.pdf";
        doc.Save(outputFile, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(outputFile))
            throw new InvalidOperationException("The expected PDF output file was not created.");
    }
}
