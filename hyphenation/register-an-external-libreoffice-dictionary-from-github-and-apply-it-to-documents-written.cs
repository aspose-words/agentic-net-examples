using System;
using System.Globalization;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a minimal LibreOffice hyphenation dictionary for Spanish (es-ES).
        const string dictionaryPath = "hyph_es_ES.dic";
        File.WriteAllText(dictionaryPath,
            "UTF-8\n" +
            "extraordinariamente=ex-tra-or-di-nar-ia-men-te\n" +
            "inconstitucional=in-cons-ti-tu-cio-nal\n" +
            "hipopotomonstrosesquipedaliofobia=hi-po-po-to-mon-stro-ses-qui-pe-da-li-o-fo-bia\n");

        // Register the dictionary with Aspose.Words.
        Hyphenation.RegisterDictionary("es-ES", dictionaryPath);

        // Create a new document and configure hyphenation options.
        Document doc = new Document();
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenationZone = 5;

        // Narrow the page width to force line wrapping and hyphenation.
        Section section = doc.FirstSection;
        section.PageSetup.PageWidth = 200; // points (~2.78 inches)
        section.PageSetup.LeftMargin = 20;
        section.PageSetup.RightMargin = 20;

        // Build content in Spanish that contains words present in the dictionary.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.LocaleId = new CultureInfo("es-ES").LCID;
        builder.Writeln("extraordinariamente inconstitucional hipopotomonstrosesquipedaliofobia");

        // Save the document as PDF.
        const string outputPath = "SpanishHyphenated.pdf";
        doc.Save(outputPath, SaveFormat.Pdf);

        // Validate that the output file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"The expected output file '{outputPath}' was not created.");
        }

        // Optional clean‑up of the temporary dictionary file.
        // File.Delete(dictionaryPath);
    }
}
