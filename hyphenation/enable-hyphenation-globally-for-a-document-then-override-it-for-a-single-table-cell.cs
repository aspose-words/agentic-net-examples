using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class HyphenationExample
{
    public static void Main()
    {
        // Prepare a minimal hyphenation dictionary for English (US).
        string dictPath = Path.Combine(Directory.GetCurrentDirectory(), "hyph_en_US.dic");
        File.WriteAllText(dictPath,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Register the dictionary so Aspose.Words can hyphenate English text.
        Hyphenation.RegisterDictionary("en-US", dictPath);
        if (!Hyphenation.IsDictionaryRegistered("en-US"))
            throw new InvalidOperationException("Failed to register the hyphenation dictionary.");

        // Create a new document and enable automatic hyphenation globally.
        Document doc = new Document();
        doc.HyphenationOptions.AutoHyphenation = true;
        // Optional: tweak hyphenation settings.
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch

        // Narrow the page width to force line wrapping and hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 200;
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a paragraph with a long word that can be hyphenated.
        builder.Writeln("extraordinarycharacteristically communication");

        // Insert a table with two cells.
        builder.StartTable();

        // First cell – hyphenation follows the global setting (enabled).
        builder.InsertCell();
        builder.Writeln("extraordinarycharacteristically");

        // Second cell – override hyphenation for this paragraph.
        builder.InsertCell();
        builder.Writeln("extraordinarycharacteristically");
        // Suppress hyphenation for the paragraph just added.
        builder.CurrentParagraph.ParagraphFormat.SuppressAutoHyphens = true;

        builder.EndTable();

        // Save the document to PDF so hyphenation can be observed.
        string outPath = Path.Combine(Directory.GetCurrentDirectory(), "HyphenationExample.pdf");
        doc.Save(outPath, SaveFormat.Pdf);

        // Verify that the output file was created.
        if (!File.Exists(outPath))
            throw new InvalidOperationException("The PDF output file was not created.");

        // Clean up the temporary dictionary file (optional).
        // File.Delete(dictPath);
    }
}
