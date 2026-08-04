using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Prepare a minimal hyphenation dictionary for English (US).
        const string dictFileName = "hyph_en_US.dic";
        File.WriteAllText(dictFileName,
            "UTF-8\n" +
            "characteristically=char-ac-ter-is-ti-cal-ly\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Register the dictionary so that Aspose.Words can hyphenate English text.
        Hyphenation.RegisterDictionary("en-US", dictFileName);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow the page width to force line wrapping and possible hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 200;
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Enable automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Set the language of the text to match the registered dictionary.
        builder.Font.LocaleId = new CultureInfo("en-US").LCID;

        // First paragraph – hyphenation disabled.
        builder.Writeln("characteristically communication");
        Paragraph firstPara = doc.FirstSection.Body.Paragraphs[0];
        firstPara.ParagraphFormat.SuppressAutoHyphens = true;

        // Target paragraph – hyphenation enabled (default).
        builder.Writeln("characteristically communication");

        // Third paragraph – hyphenation disabled.
        builder.Writeln("characteristically communication");
        Paragraph thirdPara = doc.FirstSection.Body.Paragraphs[2];
        thirdPara.ParagraphFormat.SuppressAutoHyphens = true;

        // Save the document to PDF to visualize hyphenation.
        const string outputPath = "HyphenatedParagraph.pdf";
        doc.Save(outputPath, SaveFormat.Pdf);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The expected PDF output was not created.");
    }
}
