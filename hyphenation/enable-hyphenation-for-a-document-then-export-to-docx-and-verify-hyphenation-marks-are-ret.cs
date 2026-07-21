using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a minimal hyphenation dictionary for English (US).
        const string dictFileName = "hyph_en_US.dic";
        const string dictContent =
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n";

        File.WriteAllText(dictFileName, dictContent);

        // Register the dictionary.
        Hyphenation.RegisterDictionary("en-US", dictFileName);
        if (!Hyphenation.IsDictionaryRegistered("en-US"))
            throw new InvalidOperationException("Hyphenation dictionary registration failed.");

        // Create a new document and configure a narrow page width to force wrapping.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        doc.FirstSection.PageSetup.PageWidth = 200; // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Write long words that can be hyphenated.
        builder.Font.Size = 24;
        builder.Writeln("extraordinarycharacteristically internationalization communication");

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch (in 1/20 points)
        doc.HyphenationOptions.HyphenateCaps = true;

        // Save the document to DOCX.
        const string outputFile = "Hyphenated.docx";
        doc.Save(outputFile, SaveFormat.Docx);

        // Verify that the DOCX file was created.
        if (!File.Exists(outputFile))
            throw new InvalidOperationException("The DOCX file was not created.");

        // Reload the saved document and verify hyphenation settings are retained.
        Document loaded = new Document(outputFile);
        if (!loaded.HyphenationOptions.AutoHyphenation)
            throw new InvalidOperationException("Auto hyphenation option was not retained in the saved document.");

        // Optional cleanup (commented out to keep files for inspection).
        // File.Delete(dictFileName);
        // File.Delete(outputFile);
    }
}
