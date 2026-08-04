using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class HyphenationExample
{
    public static void Main()
    {
        // Create a minimal hyphenation dictionary for en-US.
        const string dictFileName = "hyph_en_US.dic";
        File.WriteAllText(dictFileName,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Register the dictionary for the en-US language.
        Hyphenation.RegisterDictionary("en-US", dictFileName);
        if (!Hyphenation.IsDictionaryRegistered("en-US"))
            throw new InvalidOperationException("Failed to register the en-US hyphenation dictionary.");

        // Create a new document and configure page layout to force line wrapping.
        Document doc = new Document();
        doc.FirstSection.PageSetup.PageWidth = 200; // narrow width (points)
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Build content that contains words which can be hyphenated.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Size = 12;
        builder.Font.LocaleId = new CultureInfo("en-US").LCID; // set language to en-US
        builder.Writeln("extraordinarycharacteristically communication demonstration of automatic hyphenation.");

        // Enable automatic hyphenation for the document.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Save the document to PDF (layout is performed during save).
        const string outputFile = "Hyphenated.pdf";
        doc.Save(outputFile, SaveFormat.Pdf);

        // Verify that the output file was created.
        if (!File.Exists(outputFile))
            throw new InvalidOperationException($"The expected output file '{outputFile}' was not created.");

        // Additional sanity check: ensure auto hyphenation is enabled.
        if (!doc.HyphenationOptions.AutoHyphenation)
            throw new InvalidOperationException("Auto hyphenation is not enabled as expected.");

        // Clean up temporary dictionary file (optional).
        // File.Delete(dictFileName);
    }
}
