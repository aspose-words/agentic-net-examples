using System;
using System.Globalization;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow the page width to force line wrapping.
        doc.FirstSection.PageSetup.PageWidth = 200; // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Create a minimal hyphenation dictionary for English (US).
        const string dictFileName = "hyph_en_US.dic";
        File.WriteAllText(dictFileName,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Register the dictionary.
        Hyphenation.RegisterDictionary("en-US", dictFileName);

        // Enable automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;

        // First paragraph – hyphenation suppressed.
        builder.Writeln(
            "This paragraph will NOT be hyphenated: extraordinarycharacteristically internationalization communication.");

        // Second paragraph – hyphenation enabled.
        builder.Writeln(
            "This paragraph WILL be hyphenated: extraordinarycharacteristically internationalization communication.");

        // Explicitly set the hyphenation suppression flags on the paragraphs.
        Paragraph firstPara = doc.FirstSection.Body.Paragraphs[0];
        Paragraph secondPara = doc.FirstSection.Body.Paragraphs[1];
        firstPara.ParagraphFormat.SuppressAutoHyphens = true;
        secondPara.ParagraphFormat.SuppressAutoHyphens = false;

        // Save the document to PDF to visualize hyphenation.
        const string outputFile = "HyphenatedRange.pdf";
        doc.Save(outputFile, SaveFormat.Pdf);

        // Verify that the output file was created.
        if (!File.Exists(outputFile))
            throw new InvalidOperationException("Expected PDF output file was not created.");

        // Verify hyphenation settings on the two paragraphs.
        if (!firstPara.ParagraphFormat.SuppressAutoHyphens)
            throw new InvalidOperationException("First paragraph should have hyphenation suppressed.");

        if (secondPara.ParagraphFormat.SuppressAutoHyphens)
            throw new InvalidOperationException("Second paragraph should have hyphenation enabled.");

        // Clean up the temporary dictionary file.
        File.Delete(dictFileName);
    }
}
