using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a minimal hyphenation dictionary for English (US).
        const string dictionaryPath = "hyph_en_US.dic";
        string dictionaryContent =
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n";
        File.WriteAllText(dictionaryPath, dictionaryContent);

        // Register the dictionary with Aspose.Words.
        Hyphenation.RegisterDictionary("en-US", dictionaryPath);

        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Enable automatic hyphenation for the document.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Configure a narrow page width to force line wrapping.
        Section section = doc.FirstSection;
        section.PageSetup.PageWidth = 200;   // points
        section.PageSetup.LeftMargin = 20;
        section.PageSetup.RightMargin = 20;

        // Add text that contains long words defined in the dictionary.
        builder.Writeln("extraordinarycharacteristically internationalization communication");

        // Save the document to DOCX format.
        const string outputPath = "hyphenated.docx";
        doc.Save(outputPath, SaveFormat.Docx);

        // Verify that the DOCX file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The expected DOCX output file was not created.");

        // Reload the saved document and confirm hyphenation remains enabled.
        Document loadedDoc = new Document(outputPath);
        if (!loadedDoc.HyphenationOptions.AutoHyphenation)
            throw new InvalidOperationException("Hyphenation was not retained after saving the document.");

        // Optional clean‑up of the temporary dictionary file.
        // File.Delete(dictionaryPath);
    }
}
