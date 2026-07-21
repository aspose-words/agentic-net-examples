using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;
using Aspose.Words.Tables;

public class HyphenationExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow the page width so that words will need to wrap.
        doc.FirstSection.PageSetup.PageWidth = 300; // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Prepare a sample text containing long words that can be hyphenated.
        string sampleText = "extraordinarycharacteristically internationalization communication";

        // Configure paragraph formatting: larger font and increased line spacing for readability.
        builder.Font.Size = 24;
        builder.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
        builder.ParagraphFormat.LineSpacing = 18; // 1.5 lines (default line height is 12 points)

        // Write the sample text.
        builder.Writeln(sampleText);

        // Create a minimal hyphenation dictionary for English (US).
        string dictPath = Path.Combine(Directory.GetCurrentDirectory(), "hyph_en_US.dic");
        File.WriteAllText(dictPath,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Register the dictionary with Aspose.Words.
        Hyphenation.RegisterDictionary("en-US", dictPath);

        // Enable automatic hyphenation for the document.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch
        doc.HyphenationOptions.HyphenateCaps = true;

        // Save the result to PDF.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Hyphenated.pdf");
        doc.Save(outputPath, SaveFormat.Pdf);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The expected output PDF was not created.");

        // Clean up the temporary dictionary file (optional).
        if (File.Exists(dictPath))
            File.Delete(dictPath);
    }
}
