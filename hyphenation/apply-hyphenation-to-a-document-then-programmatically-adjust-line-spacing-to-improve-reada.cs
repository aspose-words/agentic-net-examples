using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Path for the hyphenation dictionary.
        const string dictFile = "hyph_en_US.dic";

        // Create a minimal hyphenation dictionary for English (US).
        // The first line must specify the encoding, followed by word=hyphenation-pattern lines.
        File.WriteAllText(dictFile,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Register the dictionary with Aspose.Words.
        Hyphenation.RegisterDictionary("en-US", dictFile);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow the page width to force line wrapping and hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 300; // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Sample text containing long words that can be hyphenated.
        string sampleText = "extraordinarycharacteristically internationalization communication " +
                            "demonstrates how automatic hyphenation can improve text layout " +
                            "when the line width is limited. Aspose.Words provides powerful " +
                            "features for document processing.";

        // Write the text into the document.
        builder.Font.Size = 12;
        builder.Writeln(sampleText);

        // Enable automatic hyphenation and configure options.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch
        doc.HyphenationOptions.HyphenateCaps = true;

        // Adjust line spacing to improve readability.
        builder.ParagraphFormat.LineSpacingRule = LineSpacingRule.Multiple;
        builder.ParagraphFormat.LineSpacing = 18; // 1.5 lines (default line height is 12 points)

        // Save the document to PDF.
        const string outputFile = "Hyphenated.pdf";
        doc.Save(outputFile, SaveFormat.Pdf);

        // Verify that the output file was created.
        if (!File.Exists(outputFile))
            throw new InvalidOperationException("The expected output file was not created.");
    }
}
