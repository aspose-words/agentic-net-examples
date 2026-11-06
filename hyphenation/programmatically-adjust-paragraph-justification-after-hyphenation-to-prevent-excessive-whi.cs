using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a minimal hyphenation dictionary for English (en-US).
        string dictionaryPath = "hyph_en_US.dic";
        string dictionaryContent =
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n";
        File.WriteAllText(dictionaryPath, dictionaryContent);

        // Register the dictionary with Aspose.Words.
        Hyphenation.RegisterDictionary("en-US", dictionaryPath);

        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Configure a narrow page width to force line wrapping and hyphenation.
        Section section = doc.FirstSection;
        PageSetup pageSetup = section.PageSetup;
        pageSetup.PageWidth = 200; // points
        pageSetup.LeftMargin = 20;
        pageSetup.RightMargin = 20;

        // Enable automatic hyphenation and set a small hyphenation zone
        // to reduce large white‑space gaps in justified text.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.HyphenationZone = 5; // points from the right margin

        // Write a paragraph with words that can be hyphenated.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Justify;
        builder.Writeln(
            "extraordinarycharacteristically internationalization communication " +
            "extraordinarycharacteristically internationalization communication");

        // Save the document as PDF.
        string outputPath = "JustifiedHyphenated.pdf";
        doc.Save(outputPath, SaveFormat.Pdf);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Expected PDF output was not created.");
    }
}
