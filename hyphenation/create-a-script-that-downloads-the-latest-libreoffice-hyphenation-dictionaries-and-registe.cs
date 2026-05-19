using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Paths for the temporary dictionary and the resulting PDF.
        const string dictionaryPath = "hyph_en_US.dic";
        const string outputPath = "hyphenated.pdf";

        // Minimal hyphenation dictionary content for English (US).
        // The first line specifies the encoding, followed by word=hyphenated-form entries.
        string dictionaryContent =
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n";

        // Write the dictionary file to the local file system.
        File.WriteAllText(dictionaryPath, dictionaryContent);

        // Register the dictionary with Aspose.Words for the "en-US" locale.
        // The Hyphenation class is in the Aspose.Words namespace.
        Hyphenation.RegisterDictionary("en-US", dictionaryPath);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Configure a narrow page width to force line wrapping and enable hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 200; // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.HyphenateCaps = true;
        doc.HyphenationOptions.HyphenationZone = 360; // default value

        // Write sample text containing long words that can be hyphenated.
        builder.Font.Size = 12;
        builder.Font.LocaleId = new CultureInfo("en-US").LCID;
        builder.Writeln("extraordinarycharacteristically internationalization communication");
        builder.Writeln("The quick brown fox jumps over the lazy dog while demonstrating hyphenation.");

        // Save the document as PDF.
        doc.Save(outputPath, SaveFormat.Pdf);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Expected output file '{outputPath}' was not created.");

        // Optional cleanup of the temporary dictionary file.
        // File.Delete(dictionaryPath);
    }
}
