using System;
using System.Globalization;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a minimal Spanish hyphenation dictionary in OpenOffice format.
        const string dictionaryFileName = "hyph_es_ES.dic";
        File.WriteAllText(dictionaryFileName,
            "UTF-8\n" +
            "extraordinariamente=ex-tra-or-di-nar-ia-men-te\n" +
            "internacionalización=in-ter-na-cio-na-li-za-ción\n" +
            "comunicación=co-mu-ni-ca-ción\n");

        // Register the dictionary for the Spanish locale.
        Hyphenation.RegisterDictionary("es-ES", dictionaryFileName);
        if (!Hyphenation.IsDictionaryRegistered("es-ES"))
            throw new InvalidOperationException("Failed to register the Spanish hyphenation dictionary.");

        // Build a document containing Spanish text that can be hyphenated.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set a narrow page width to force line wrapping.
        doc.FirstSection.PageSetup.PageWidth = 200; // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Use a Spanish locale for the text.
        builder.Font.LocaleId = new CultureInfo("es-ES").LCID;
        builder.Font.Size = 12;

        // Write a paragraph with words that have hyphenation patterns.
        builder.Writeln("extraordinariamente internacionalización comunicación");

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Save the document to PDF to render hyphenation.
        const string outputFile = "HyphenatedSpanish.pdf";
        doc.Save(outputFile);

        // Verify that the output file was created.
        if (!File.Exists(outputFile))
            throw new InvalidOperationException("The expected PDF output file was not created.");
    }
}
