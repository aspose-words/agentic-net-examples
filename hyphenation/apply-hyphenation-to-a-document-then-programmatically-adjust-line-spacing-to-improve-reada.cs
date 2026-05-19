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

        // Set a narrow page width to force line wrapping and hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 300; // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Prepare a minimal hyphenation dictionary for English (US).
        const string dictFileName = "hyph_en_US.dic";
        File.WriteAllText(dictFileName,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Register the dictionary.
        Hyphenation.RegisterDictionary("en-US", dictFileName);

        // Write sample text that contains long words.
        builder.Font.Size = 12;
        builder.Font.LocaleId = new CultureInfo("en-US").LCID;
        builder.Writeln("extraordinarycharacteristically internationalization communication");

        // Enable automatic hyphenation and adjust its options.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch
        doc.HyphenationOptions.HyphenateCaps = true;
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;

        // Adjust line spacing to improve readability.
        builder.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
        builder.ParagraphFormat.LineSpacing = 18; // 1.5 lines (12 * 1.5)

        // Save the document as PDF.
        const string outputFile = "Hyphenated.pdf";
        doc.Save(outputFile, SaveFormat.Pdf);

        // Validate that the output file was created.
        if (!File.Exists(outputFile))
            throw new InvalidOperationException($"Expected output file '{outputFile}' was not created.");

        // Clean up the temporary dictionary file.
        if (File.Exists(dictFileName))
            File.Delete(dictFileName);
    }
}
