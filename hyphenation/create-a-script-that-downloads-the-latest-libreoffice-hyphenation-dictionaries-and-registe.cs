using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Paths for the temporary hyphenation dictionary and the resulting PDF.
        const string dictionaryPath = "hyph_en_US.dic";
        const string outputPath = "Hyphenated.pdf";

        // Minimal hyphenation dictionary content (OpenOffice format).
        string dictionaryContent =
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n";

        // Write the dictionary file locally.
        File.WriteAllText(dictionaryPath, dictionaryContent);

        // Verify the dictionary file exists.
        if (!File.Exists(dictionaryPath))
            throw new InvalidOperationException($"Dictionary file '{dictionaryPath}' was not created.");

        // Register the dictionary for the "en-US" locale.
        Hyphenation.RegisterDictionary("en-US", dictionaryPath);

        // Ensure registration succeeded.
        if (!Hyphenation.IsDictionaryRegistered("en-US"))
            throw new InvalidOperationException("Failed to register the hyphenation dictionary.");

        // Create a new blank document and add sample text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("extraordinarycharacteristically internationalization communication");

        // Narrow page width to force line wrapping and enable automatic hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 200;
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;
        doc.HyphenationOptions.AutoHyphenation = true;

        // Save the document as PDF.
        doc.Save(outputPath, SaveFormat.Pdf);

        // Verify the PDF was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Output file '{outputPath}' was not created.");

        // Optional cleanup of the temporary dictionary file.
        // File.Delete(dictionaryPath);
    }
}
