using System;
using System.Globalization;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a minimal LibreOffice‑style Spanish hyphenation dictionary.
        const string dictionaryPath = "hyph_es_ES.dic";
        File.WriteAllText(dictionaryPath,
            "UTF-8\n" +
            "extraordinariamente=ex-tra-or-di-nar-ia-men-te\n" +
            "inconstitucionalmente=in-cons-ti-tu-cio-nal-men-te\n" +
            "electroencefalografista=elec-tro-en-ce-fa-lo-gra-fis-ta\n");

        // Register the dictionary for the Spanish locale (es-ES).
        Aspose.Words.Hyphenation.RegisterDictionary("es-ES", dictionaryPath);

        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the paragraph (font) locale to Spanish so hyphenation uses the registered dictionary.
        builder.Font.LocaleId = new CultureInfo("es-ES").LCID;

        // Write a line of Spanish words that are long enough to require hyphenation.
        builder.Writeln("extraordinariamente inconstitucionalmente electroencefalografista");

        // Narrow the page width to force line wrapping and make hyphenation visible.
        Section section = doc.FirstSection;
        section.PageSetup.PageWidth = 200; // points
        section.PageSetup.LeftMargin = 20;
        section.PageSetup.RightMargin = 20;
        section.PageSetup.TopMargin = 20;
        section.PageSetup.BottomMargin = 20;

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.HyphenationZone = 5;

        // Save the document as PDF to demonstrate hyphenation.
        const string outputPath = "HyphenatedSpanish.pdf";
        doc.Save(outputPath);

        // Verify that the PDF was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("Expected PDF output was not created.");
        }
    }
}
