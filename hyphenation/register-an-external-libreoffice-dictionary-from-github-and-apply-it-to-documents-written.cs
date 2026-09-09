using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Paths for the dictionary file and the output PDF.
        const string dictionaryPath = "hyph_es_ES.dic";
        const string outputPdfPath = "HyphenatedSpanish.pdf";

        // Create a minimal Spanish hyphenation dictionary in OpenOffice format.
        // First line must specify the encoding, followed by word=hyphenation patterns.
        string dictionaryContent =
            "UTF-8\n" +
            "extraordinariamente=ex-tra-or-di-nar-ia-men-te\n" +
            "hipopotomonstrosesquipedaliofobia=hi-po-po-to-mo-ns-tro-se-squi-pe-da-li-o-fo-bia\n" +
            "desafortunadamente=de-sa-for-tu-na-da-men-te\n";

        File.WriteAllText(dictionaryPath, dictionaryContent);

        // Register the dictionary for the Spanish (Spain) locale.
        Hyphenation.RegisterDictionary("es-ES", dictionaryPath);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Configure page layout to force line wrapping.
        doc.FirstSection.PageSetup.PageWidth = 300; // points (~4.2 inches)
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Write Spanish text containing long words that can be hyphenated.
        builder.Font.Size = 24;
        builder.Font.LocaleId = new CultureInfo("es-ES").LCID;
        builder.Writeln(
            "Esta es una demostración de hyphenation automática con palabras como extraordinariamente, " +
            "hipopotomonstrosesquipedaliofobia y desafortunadamente para observar cómo se insertan guiones.");

        // Save the document as PDF.
        doc.Save(outputPdfPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(outputPdfPath))
            throw new InvalidOperationException("The expected PDF file was not created.");
    }
}
