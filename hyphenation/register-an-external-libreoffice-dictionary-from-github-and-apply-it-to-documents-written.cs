using System;
using System.Globalization;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a minimal Spanish hyphenation dictionary in OpenOffice format.
        const string dictFileName = "hyph_es_ES.dic";
        string dictContent =
            "UTF-8\n" +
            "extraordinario=ex-tra-or-di-nar-io\n" +
            "internacionalización=in-ter-na-cio-na-li-za-cion\n" +
            "comunicación=co-mu-ni-ca-ción\n";

        File.WriteAllText(dictFileName, dictContent);

        // Register the dictionary for the Spanish locale.
        Aspose.Words.Hyphenation.RegisterDictionary("es-ES", dictFileName);
        if (!Aspose.Words.Hyphenation.IsDictionaryRegistered("es-ES"))
            throw new InvalidOperationException("Spanish hyphenation dictionary was not registered.");

        // Build a document containing Spanish text that can be hyphenated.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set a narrow page width to force line wrapping.
        doc.FirstSection.PageSetup.PageWidth = 200;
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Set the locale of the text to Spanish (Spain).
        builder.Font.LocaleId = new CultureInfo("es-ES").LCID;

        // Write a paragraph with words that match the dictionary entries.
        builder.Writeln(
            "extraordinario internacionalización comunicación " +
            "extraordinario internacionalización comunicación " +
            "extraordinario internacionalización comunicación");

        // Save the document as PDF to see hyphenation in effect.
        const string outputFile = "HyphenatedSpanish.pdf";
        doc.Save(outputFile);

        // Verify that the output file was created.
        if (!File.Exists(outputFile))
            throw new InvalidOperationException("Expected output PDF was not created.");
    }
}
