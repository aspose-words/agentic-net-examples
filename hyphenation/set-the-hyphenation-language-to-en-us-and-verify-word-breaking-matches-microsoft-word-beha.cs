using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a long word that can be hyphenated.
        builder.Font.Size = 24;
        builder.Font.LocaleId = new CultureInfo("en-US").LCID;
        builder.Writeln("extraordinarycharacteristically");

        // Narrow the page width to force line wrapping and hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 200;
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Create a minimal hyphenation dictionary for en‑US.
        const string dictFileName = "hyph_en_US.dic";
        File.WriteAllText(dictFileName,
            "UTF-8\nextraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n");

        // Register the dictionary.
        Hyphenation.RegisterDictionary("en-US", dictFileName);

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Save the document to PDF (fixed‑page format where hyphenation is applied).
        const string outputFile = "Hyphenated.pdf";
        doc.Save(outputFile, SaveFormat.Pdf);

        // Verify that the output file was created.
        if (!File.Exists(outputFile))
            throw new InvalidOperationException("The hyphenated PDF was not created.");

        // Optional sanity check: ensure the dictionary is registered.
        if (!Hyphenation.IsDictionaryRegistered("en-US"))
            throw new InvalidOperationException("The en‑US hyphenation dictionary is not registered.");
    }
}
