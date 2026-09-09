using System;
using System.Globalization;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Paths for the dictionary and the output PDF.
        const string dictionaryPath = "hyph_en_US.dic";
        const string outputPath = "hyphenated.pdf";

        // Create a minimal hyphenation dictionary for English (US).
        // The first line must be the encoding identifier.
        // Subsequent lines contain word=hyphenation-patterns.
        string dictionaryContent =
            "UTF-8\n" +
            "extraordinarycharacteristically=ex-tra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n";

        File.WriteAllText(dictionaryPath, dictionaryContent);

        // Register the dictionary with Aspose.Words.
        Hyphenation.RegisterDictionary("en-US", dictionaryPath);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the document locale to English (US) so that the registered dictionary is used.
        builder.Font.LocaleId = new CultureInfo("en-US").LCID;

        // Write a paragraph containing long words that can be hyphenated.
        builder.Writeln(
            "extraordinarycharacteristically internationalization communication " +
            "extraordinarycharacteristically internationalization communication");

        // Narrow the page width to force line wrapping and enable hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 200; // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.HyphenationZone = 360; // default
        doc.HyphenationOptions.HyphenateCaps = true;
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;

        // Save the document as PDF.
        doc.Save(outputPath, SaveFormat.Pdf);

        // Validate that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Expected output file '{outputPath}' was not created.");

        // Clean up temporary dictionary file (optional).
        // File.Delete(dictionaryPath);
    }
}
