using System;
using System.Globalization;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a minimal Spanish hyphenation dictionary in OpenOffice format.
        const string dictionaryPath = "hyph_es_ES.dic";
        File.WriteAllText(dictionaryPath,
            "UTF-8\n" +
            "extraordinariamente=ex-tra-or-di-nar-ia-men-te\n" +
            "hipopotomonstrosesquipedaliofobia=hi-po-po-to-mons-tro-se-squi-pe-da-li-o-fo-bia\n");

        // Register the dictionary for the Spanish locale.
        // Note: Use the static Hyphenation class directly (no extra nesting).
        Aspose.Words.Hyphenation.RegisterDictionary("es-ES", dictionaryPath);

        // Build a document with Spanish text that contains long words.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the paragraph locale to Spanish.
        builder.Font.LocaleId = new CultureInfo("es-ES").LCID;

        // Write sample text.
        builder.Writeln(
            "La palabra extraordinariamente es muy larga y puede ser dividida. " +
            "Otro término complicado es hipopotomonstrosesquipedaliofobia, que también necesita hyphenation.");

        // Narrow the page to force line wrapping and hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 200;
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Save the document as PDF.
        const string outputPath = "hyphenated_es.pdf";
        doc.Save(outputPath);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The expected PDF output was not created.");
    }
}
