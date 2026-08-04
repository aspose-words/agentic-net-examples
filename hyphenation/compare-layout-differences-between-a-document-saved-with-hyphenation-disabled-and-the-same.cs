using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Prepare a minimal hyphenation dictionary for English (en-US).
        string dictPath = Path.Combine(Directory.GetCurrentDirectory(), "hyph_en_US.dic");
        File.WriteAllText(dictPath,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Register the dictionary so that Aspose.Words can hyphenate the words.
        Hyphenation.RegisterDictionary("en-US", dictPath);

        // Create and save a document with hyphenation disabled.
        Document docNoHyphen = CreateDocument(autoHyphenation: false);
        string noHyphenPath = "HyphenationDisabled.pdf";
        docNoHyphen.Save(noHyphenPath);
        if (!File.Exists(noHyphenPath))
            throw new InvalidOperationException("Failed to create the hyphenation‑disabled PDF.");

        // Create and save a document with hyphenation enabled.
        Document docHyphen = CreateDocument(autoHyphenation: true);
        string hyphenPath = "HyphenationEnabled.pdf";
        docHyphen.Save(hyphenPath);
        if (!File.Exists(hyphenPath))
            throw new InvalidOperationException("Failed to create the hyphenation‑enabled PDF.");

        // Compare the resulting files (e.g., by file size) to illustrate a layout difference.
        long sizeNoHyphen = new FileInfo(noHyphenPath).Length;
        long sizeHyphen = new FileInfo(hyphenPath).Length;

        Console.WriteLine($"Hyphenation disabled PDF size: {sizeNoHyphen} bytes");
        Console.WriteLine($"Hyphenation enabled PDF size: {sizeHyphen} bytes");
        Console.WriteLine(sizeHyphen != sizeNoHyphen
            ? "The PDFs differ, indicating a layout change due to hyphenation."
            : "The PDFs are identical; hyphenation did not affect the layout.");
    }

    // Creates a document containing long words that can be hyphenated.
    private static Document CreateDocument(bool autoHyphenation)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow the page width to force line wrapping and potential hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 200; // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Use a font size that makes the text long enough to wrap.
        builder.Font.Size = 12;
        builder.Font.LocaleId = new CultureInfo("en-US").LCID;

        // Write a paragraph with words that have hyphenation points defined in the dictionary.
        builder.Writeln("extraordinarycharacteristically internationalization communication");

        // Configure hyphenation options for this document.
        doc.HyphenationOptions.AutoHyphenation = autoHyphenation;
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch
        doc.HyphenationOptions.HyphenateCaps = true;

        // Ensure the layout is up‑to‑date before saving.
        doc.UpdatePageLayout();

        return doc;
    }
}
