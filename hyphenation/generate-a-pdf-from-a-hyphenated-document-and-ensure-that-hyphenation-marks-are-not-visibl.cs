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

        // Write a paragraph containing long words that can be hyphenated.
        builder.Writeln(
            "extraordinarycharacteristically internationalization communication " +
            "extraordinarycharacteristically internationalization communication " +
            "extraordinarycharacteristically internationalization communication");

        // Narrow the page width to force line wrapping and hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 200; // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Create a minimal hyphenation dictionary for English (US) in the current folder.
        string dictFileName = Path.Combine(Directory.GetCurrentDirectory(), "hyph_en_US.dic");
        string dictContent =
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n";

        File.WriteAllText(dictFileName, dictContent);

        // Register the dictionary.
        Hyphenation.RegisterDictionary("en-US", dictFileName);

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Optional: set a valid hyphenation zone (default is 360). Setting to 0 throws an exception.
        doc.HyphenationOptions.HyphenationZone = 360; // 0.25 inch from the right margin

        // Save the document as PDF. Hyphenation marks (soft hyphens) are not inserted manually,
        // so they will not appear as visible characters in the output PDF.
        const string outputPdf = "hyphenated.pdf";
        doc.Save(outputPdf);

        // Validate that the PDF was created.
        if (!File.Exists(outputPdf))
            throw new InvalidOperationException("Expected PDF output file was not created.");
    }
}
