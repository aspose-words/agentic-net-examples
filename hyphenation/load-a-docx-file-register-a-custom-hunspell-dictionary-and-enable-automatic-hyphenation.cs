using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using static Aspose.Words.Hyphenation;

public class HyphenationExample
{
    public static void Main()
    {
        // File names used in the example.
        const string inputDocx = "sample.docx";
        const string dictionaryFile = "hyph_en_US.dic";
        const string outputPdf = "hyphenated.pdf";

        // Create a document with long words that can be hyphenated.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Size = 24;
        builder.Font.LocaleId = new CultureInfo("en-US").LCID;
        builder.Writeln("extraordinarycharacteristically internationalization communication");

        // Narrow the page width to force line wrapping and hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 200;
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Save the source document as DOCX.
        doc.Save(inputDocx);

        // Create a minimal Hunspell dictionary file for English (US).
        string dictionaryContent =
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n";
        File.WriteAllText(dictionaryFile, dictionaryContent);

        // Load the previously saved DOCX document.
        Document loadedDoc = new Document(inputDocx);

        // Register the custom hyphenation dictionary.
        RegisterDictionary("en-US", dictionaryFile);

        // Enable automatic hyphenation for the document.
        loadedDoc.HyphenationOptions.AutoHyphenation = true;

        // Save the hyphenated document as PDF.
        loadedDoc.Save(outputPdf);

        // Verify that the output file was created.
        if (!File.Exists(outputPdf))
            throw new InvalidOperationException("The hyphenated PDF was not created.");
    }
}
