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
        doc.FirstSection.PageSetup.PageWidth = 200;
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

        // Enable automatic hyphenation for the document.
        doc.HyphenationOptions.AutoHyphenation = true;

        // First paragraph – hyphenation disabled.
        builder.Font.LocaleId = new CultureInfo("en-US").LCID;
        builder.Writeln("extraordinarycharacteristically internationalization communication");
        Paragraph firstPara = doc.FirstSection.Body.Paragraphs[0];
        firstPara.ParagraphFormat.SuppressAutoHyphens = true;

        // Second paragraph – hyphenation enabled (default).
        builder.Writeln("extraordinarycharacteristically internationalization communication");

        // Third paragraph – hyphenation disabled.
        builder.Writeln("extraordinarycharacteristically internationalization communication");
        Paragraph thirdPara = doc.FirstSection.Body.Paragraphs[2];
        thirdPara.ParagraphFormat.SuppressAutoHyphens = true;

        // Save the document to PDF to visualize hyphenation.
        const string outputFile = "Hyphenated.pdf";
        doc.Save(outputFile, SaveFormat.Pdf);

        // Validate that the output file was created.
        if (!File.Exists(outputFile))
            throw new InvalidOperationException($"Expected output file '{outputFile}' was not created.");
    }
}
