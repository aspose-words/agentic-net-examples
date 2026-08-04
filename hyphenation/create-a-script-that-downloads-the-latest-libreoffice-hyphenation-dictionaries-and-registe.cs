using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Define file names.
        const string dictionaryFileName = "hyph_en_US.dic";
        const string outputFileName = "hyphenated.pdf";

        // Create a minimal hyphenation dictionary for English (US).
        // The first line must specify the encoding, followed by word=hyphenation patterns.
        string dictionaryContent =
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n";

        File.WriteAllText(dictionaryFileName, dictionaryContent);

        // Register the dictionary with Aspose.Words.
        Hyphenation.RegisterDictionary("en-US", dictionaryFileName);

        // Verify registration.
        if (!Hyphenation.IsDictionaryRegistered("en-US"))
            throw new InvalidOperationException("Failed to register the hyphenation dictionary.");

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set a narrow page width to force line wrapping and hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 200; // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Write sample text containing long words that can be hyphenated.
        builder.Font.Size = 12;
        builder.Font.LocaleId = new CultureInfo("en-US").LCID;
        builder.Writeln("extraordinarycharacteristically internationalization communication");

        // Save the document to PDF (any format that triggers layout).
        doc.Save(outputFileName, SaveFormat.Pdf);

        // Validate that the output file was created.
        if (!File.Exists(outputFileName))
            throw new InvalidOperationException($"The expected output file '{outputFileName}' was not created.");

        // Clean up temporary dictionary file (optional).
        // File.Delete(dictionaryFileName);
    }
}
