using System;
using System.Globalization;
using System.IO;
using Aspose.Words;

public class HyphenationExample
{
    public static void Main()
    {
        // Create a minimal hyphenation dictionary for English (US).
        const string dictFileName = "hyph_en_US.dic";
        File.WriteAllText(dictFileName,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "communication=com-mu-ni-ca-tion\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n");

        // Register the dictionary so that Aspose.Words can hyphenate words of this language.
        Hyphenation.RegisterDictionary("en-US", dictFileName);

        // -----------------------------------------------------------------
        // Create a source document that contains long words which can be hyphenated.
        // -----------------------------------------------------------------
        const string sourceFile = "source.docx";
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Narrow the page width to force line wrapping and thus hyphenation.
        sourceDoc.FirstSection.PageSetup.PageWidth = 200; // points
        sourceDoc.FirstSection.PageSetup.LeftMargin = 20;
        sourceDoc.FirstSection.PageSetup.RightMargin = 20;

        // Set the language of the text to English (US) so the registered dictionary is used.
        builder.Font.LocaleId = new CultureInfo("en-US").LCID;
        builder.Font.Size = 12;

        // Write a paragraph with words that have hyphenation patterns defined above.
        builder.Writeln(
            "extraordinarycharacteristically communication internationalization " +
            "extraordinarycharacteristically communication internationalization");

        // Save the source document.
        sourceDoc.Save(sourceFile);
        if (!File.Exists(sourceFile))
            throw new InvalidOperationException("Source DOCX was not created.");

        // -----------------------------------------------------------------
        // Load the DOCX, enable automatic hyphenation, and save the result.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(sourceFile);

        // Enable automatic hyphenation for the document.
        loadedDoc.HyphenationOptions.AutoHyphenation = true;
        // Optional: configure additional hyphenation options.
        loadedDoc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        loadedDoc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch
        loadedDoc.HyphenationOptions.HyphenateCaps = true;

        // Force layout update so hyphenation is applied before saving.
        loadedDoc.UpdatePageLayout();

        const string outputFile = "hyphenated.docx";
        loadedDoc.Save(outputFile);
        if (!File.Exists(outputFile))
            throw new InvalidOperationException("Hyphenated DOCX was not created.");
    }
}
