using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Create a minimal Arabic hyphenation dictionary locally.
        const string dictFile = "hyph_ar_SA.dic";
        File.WriteAllText(dictFile,
            "UTF-8\n" +
            "مثال=مث-ال\n" +
            "تجربة=تج-ري-بة\n");

        // Register the dictionary for the Arabic (Saudi Arabia) locale.
        Hyphenation.RegisterDictionary("ar-SA", dictFile);

        if (!Hyphenation.IsDictionaryRegistered("ar-SA"))
            throw new InvalidOperationException("Arabic hyphenation dictionary registration failed.");

        // Build a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow the page width to force line wrapping.
        doc.FirstSection.PageSetup.PageWidth = 300; // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Set a valid hyphenation zone (default is 360 = 0.25 inch).
        // Using 0 throws an ArgumentOutOfRangeException, so we keep the default.
        doc.HyphenationOptions.HyphenationZone = 360;

        // Configure the paragraph for right‑to‑left layout and Arabic locale.
        builder.ParagraphFormat.Bidi = true;
        builder.Font.LocaleId = new CultureInfo("ar-SA").LCID;

        // Add Arabic text that contains words defined in the dictionary.
        builder.Font.Name = "Arial";
        builder.Font.Size = 24;
        builder.Writeln(
            "هذه جملة طويلة تحتوي على مثال وتجربة لتوضيح كيفية تطبيق الفواصل في النص العربي عندما يكون هناك حاجة لتقسيم الكلمات عبر السطر.");

        // Save the document as PDF.
        const string outFile = "HyphenatedArabic.pdf";
        doc.Save(outFile, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(outFile))
            throw new InvalidOperationException("The PDF output file was not created.");
    }
}
