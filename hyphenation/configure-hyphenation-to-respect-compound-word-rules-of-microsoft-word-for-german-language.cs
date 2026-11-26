using System;
using System.Globalization;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a minimal German hyphenation dictionary containing a compound word.
        const string dictionaryPath = "hyph_de_DE.dic";
        File.WriteAllText(dictionaryPath,
            "UTF-8\n" +
            "Donaudampfschifffahrtsgesellschaft=Do-nau-dampf-schiff-fahrts-gesell-schaft\n");

        // Register the German dictionary for hyphenation.
        Aspose.Words.Hyphenation.RegisterDictionary("de-DE", dictionaryPath);

        // Create a new document and set up a narrow page width to force line wrapping.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the document locale to German (de-DE) so that hyphenation uses the German rules.
        builder.Font.LocaleId = new CultureInfo("de-DE").LCID;

        // Configure page layout.
        Section section = doc.FirstSection;
        section.PageSetup.PageWidth = 200;   // points
        section.PageSetup.LeftMargin = 20;
        section.PageSetup.RightMargin = 20;

        // Add German text containing the compound word that can be hyphenated.
        builder.Writeln(
            "Die Donaudampfschifffahrtsgesellschaft ist ein berühmtes Beispiel für ein langes deutsches Kompositum.");

        // Save the document as PDF to visualize hyphenation.
        const string outputPath = "GermanHyphenated.pdf";
        doc.Save(outputPath, SaveFormat.Pdf);

        // Validate that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Expected PDF output was not created.");
    }
}
