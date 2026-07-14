using System;
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

        // Narrow the page width to force line wrapping and hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 200;
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Create a minimal hyphenation dictionary for English (US).
        const string dictPath = "hyph_en_US.dic";
        File.WriteAllText(dictPath,
            "UTF-8\n" +
            "extraordinarycharacteristically=ex-tra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "communication=com-mu-ni-ca-tion\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n");

        // Register the dictionary.
        Hyphenation.RegisterDictionary("en-US", dictPath);
        if (!Hyphenation.IsDictionaryRegistered("en-US"))
            throw new InvalidOperationException("Hyphenation dictionary was not registered.");

        // Enable automatic hyphenation for the document.
        doc.HyphenationOptions.AutoHyphenation = true;

        // First paragraph – hyphenation disabled.
        builder.Writeln("This paragraph will not be hyphenated even though it contains extraordinarycharacteristically long words.");
        doc.FirstSection.Body.Paragraphs[0].ParagraphFormat.SuppressAutoHyphens = true;

        // Second paragraph – hyphenation enabled.
        builder.Writeln("This paragraph will be hyphenated with extraordinarycharacteristically long words.");
        // No need to change SuppressAutoHyphens; default is false.

        // Third paragraph – hyphenation disabled.
        builder.Writeln("Again, this paragraph will not be hyphenated despite containing extraordinarycharacteristically words.");
        doc.FirstSection.Body.Paragraphs[2].ParagraphFormat.SuppressAutoHyphens = true;

        // Save the document to PDF.
        const string outputPath = "HyphenationExample.pdf";
        doc.Save(outputPath, SaveFormat.Pdf);

        // Validate that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The expected PDF output was not created.");
    }
}
