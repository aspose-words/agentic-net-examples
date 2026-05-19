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
            "communication=com-mu-ni-ca-tion\n");

        // Register the dictionary so that hyphenation can be applied.
        Hyphenation.RegisterDictionary("en-US", dictFile);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow the page width to force line wrapping and hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 200;   // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // First paragraph – hyphenation disabled.
        builder.Writeln("extraordinarycharacteristically communication");
        doc.FirstSection.Body.Paragraphs[0].ParagraphFormat.SuppressAutoHyphens = true;

        // Second paragraph – hyphenation enabled (default).
        builder.Writeln("extraordinarycharacteristically communication");
        // Explicitly ensure hyphenation is not suppressed.
        doc.FirstSection.Body.Paragraphs[1].ParagraphFormat.SuppressAutoHyphens = false;

        // Third paragraph – hyphenation disabled.
        builder.Writeln("extraordinarycharacteristically communication");
        doc.FirstSection.Body.Paragraphs[2].ParagraphFormat.SuppressAutoHyphens = true;

        // Turn on automatic hyphenation for the document.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Save the result to PDF.
        const string outFile = "Hyphenated.pdf";
        doc.Save(outFile, SaveFormat.Pdf);

        // Verify that the output file was created.
        if (!File.Exists(outFile))
            throw new InvalidOperationException($"The file '{outFile}' was not created.");
    }
}
