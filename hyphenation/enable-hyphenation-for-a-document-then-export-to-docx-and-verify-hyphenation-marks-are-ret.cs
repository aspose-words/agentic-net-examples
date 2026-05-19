using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Prepare a minimal hyphenation dictionary for English (US).
        const string dictFileName = "hyph_en_US.dic";
        const string dictContent = "UTF-8\nextraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\ninternationalization=in-ter-na-tion-al-i-za-tion\ncommunication=com-mu-ni-ca-tion\n";
        File.WriteAllText(dictFileName, dictContent);

        // Register the dictionary.
        Hyphenation.RegisterDictionary("en-US", dictFileName);

        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow the page to force line wrapping and hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 200;
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Write a long word that can be hyphenated.
        builder.Font.Size = 24;
        builder.Writeln("extraordinarycharacteristically internationalization communication");

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenationZone = 720;
        doc.HyphenationOptions.HyphenateCaps = true;

        // Save the document to DOCX.
        const string outputFile = "Hyphenated.docx";
        doc.Save(outputFile, SaveFormat.Docx);

        // Verify that the output file was created.
        if (!File.Exists(outputFile))
            throw new InvalidOperationException("The DOCX file was not created.");

        // Load the saved document and verify hyphenation settings are retained.
        Document loaded = new Document(outputFile);
        if (!loaded.HyphenationOptions.AutoHyphenation)
            throw new InvalidOperationException("Hyphenation was not retained after saving.");

        // Clean up temporary files (optional).
        // File.Delete(dictFileName);
        // File.Delete(outputFile);
    }
}
