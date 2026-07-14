using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Create a minimal hyphenation dictionary for English (US).
        const string dictFile = "hyph_en_US.dic";
        File.WriteAllText(dictFile,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Register the dictionary so that Aspose.Words can hyphenate the words.
        Hyphenation.RegisterDictionary("en-US", dictFile);

        // Create a new document and configure page layout to force line wrapping.
        Document doc = new Document();
        doc.FirstSection.PageSetup.PageWidth = 300; // narrow width
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Build content with long words that can be hyphenated.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Size = 12;
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Justify;
        builder.Writeln(
            "extraordinarycharacteristically internationalization communication " +
            "extraordinarycharacteristically internationalization communication.");

        // Enable automatic hyphenation and adjust its behavior.
        doc.HyphenationOptions.AutoHyphenation = true;
        // Use a positive value for HyphenationZone (default is 360 = 0.25 inch).
        doc.HyphenationOptions.HyphenationZone = 360;
        doc.HyphenationOptions.HyphenateCaps = true;

        // Compress spacing for justified paragraphs to avoid large gaps.
        doc.JustificationMode = JustificationMode.Compress;

        // Save the result.
        const string outputFile = "AdjustedHyphenation.pdf";
        doc.Save(outputFile, SaveFormat.Pdf);

        // Verify that the output file was created.
        if (!File.Exists(outputFile))
            throw new InvalidOperationException("The expected output file was not created.");
    }
}
