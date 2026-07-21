using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Create a blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow the page so that words are forced to wrap and hyphenate.
        doc.FirstSection.PageSetup.PageWidth = 200; // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Write a paragraph containing a long word (will be hyphenated) and a short word (should not be hyphenated).
        builder.Font.Size = 12;
        builder.Writeln("extraordinarycharacteristically test hyphenation");

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Note: Aspose.Words does not expose a MinimumWordLength property.
        // The effect of a minimum word length can be observed by the fact that short words (e.g., "test")
        // are not hyphenated while longer words are.

        // Create a minimal hyphenation dictionary for English (US).
        const string dictFileName = "hyph_en_US.dic";
        File.WriteAllText(dictFileName,
            "UTF-8\n" +
            "extraordinarycharacteristically=ex-tra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "hyphenation=hy-phen-a-tion\n");

        // Register the dictionary.
        Hyphenation.RegisterDictionary("en-US", dictFileName);

        // Save the document to PDF to observe hyphenation.
        const string outFile = "HyphenationMinWordLength.pdf";
        doc.Save(outFile, SaveFormat.Pdf);

        // Verify that the output file was created.
        if (!File.Exists(outFile))
            throw new InvalidOperationException("The PDF file was not created.");

        // Clean up temporary dictionary file.
        File.Delete(dictFileName);
    }
}
