using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class HyphenationAndLineSpacingExample
{
    public static void Main()
    {
        // Path for the temporary hyphenation dictionary.
        const string dictFileName = "hyph_en_US.dic";

        // Create a minimal hyphenation dictionary for English (US).
        // The first line must be the encoding, followed by word=hyphenation-pattern lines.
        File.WriteAllText(dictFileName,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Register the dictionary so that Aspose.Words can hyphenate the words above.
        Hyphenation.RegisterDictionary("en-US", dictFileName);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow the page width to force line wrapping where hyphenation can occur.
        doc.FirstSection.PageSetup.PageWidth = 300; // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Write a paragraph containing long words that match the dictionary entries.
        builder.Font.Size = 12;
        builder.Writeln("extraordinarycharacteristically internationalization communication");

        // Enable automatic hyphenation and configure its options.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenateCaps = true;

        // Adjust line spacing to improve readability.
        builder.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
        builder.ParagraphFormat.LineSpacing = 18; // 1.5 lines (default line height is 12 points)

        // Save the document to PDF so that hyphenation and spacing are visible.
        const string outputFile = "HyphenatedAndSpaced.pdf";
        doc.Save(outputFile, SaveFormat.Pdf);

        // Verify that the output file was created.
        if (!File.Exists(outputFile))
            throw new InvalidOperationException("The output PDF was not created.");

        // Clean up the temporary dictionary file.
        if (File.Exists(dictFileName))
            File.Delete(dictFileName);
    }
}
