using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Create a minimal hyphenation dictionary for en‑US.
        const string dictFile = "hyph_en_US.dic";
        File.WriteAllText(dictFile,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Register the dictionary.
        Hyphenation.RegisterDictionary("en-US", dictFile);
        if (!Hyphenation.IsDictionaryRegistered("en-US"))
            throw new InvalidOperationException("Failed to register the en‑US hyphenation dictionary.");

        // Create a document with narrow page width to force line breaks.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        doc.FirstSection.PageSetup.PageWidth = 200;   // 200 points (~2.78 inches)
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Set the locale to en‑US and write long words that can be hyphenated.
        builder.Font.Size = 24;
        builder.Font.LocaleId = new CultureInfo("en-US").LCID;
        builder.Writeln("extraordinarycharacteristically internationalization communication");

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Save the document to PDF (layout is performed during save).
        const string outFile = "Hyphenated.pdf";
        doc.Save(outFile, SaveFormat.Pdf);

        // Verify that the output file was created.
        if (!File.Exists(outFile))
            throw new InvalidOperationException($"The expected output file '{outFile}' was not created.");

        // Clean up temporary files (optional).
        // File.Delete(dictFile);
    }
}
