using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class HyphenationJustificationExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow the page width to force line wrapping and hyphenation.
        Section section = doc.FirstSection;
        section.PageSetup.PageWidth = 200; // points
        section.PageSetup.LeftMargin = 20;
        section.PageSetup.RightMargin = 20;

        // Create a minimal hyphenation dictionary for English (en-US).
        // The dictionary format: first line is the encoding, subsequent lines are word=hyphenation points.
        string dictPath = "hyph_en_US.dic";
        File.WriteAllText(dictPath,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Register the dictionary with Aspose.Words.
        Hyphenation.RegisterDictionary("en-US", dictPath);

        // Enable automatic hyphenation for the document.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.HyphenationZone = 5; // percent of line width
        // LanguageId property is not required for recent versions; the default language will be used.

        // Write a paragraph that contains long words which will be hyphenated.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Justify; // initial justification
        builder.Writeln("extraordinarycharacteristically internationalization communication");

        // After hyphenation, adjust justification to reduce large gaps.
        // If JustifyLow is unavailable, fall back to regular Justify.
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            para.ParagraphFormat.Alignment = ParagraphAlignment.Justify;
        }

        // Save the document as PDF to visualize the effect.
        string outputPath = "HyphenatedJustified.pdf";
        doc.Save(outputPath, SaveFormat.Pdf);

        // Validate that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The PDF output file was not created.");

        // Clean up the temporary dictionary file.
        if (File.Exists(dictPath))
            File.Delete(dictPath);
    }
}
