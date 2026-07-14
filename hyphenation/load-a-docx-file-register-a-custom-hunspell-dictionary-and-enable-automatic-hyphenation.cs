using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX file with long text that can be hyphenated.
        const string inputPath = "sample.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Size = 24;
        builder.Writeln("extraordinarycharacteristically internationalization communication");
        // Narrow the page to force line wrapping and hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 200;
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;
        doc.Save(inputPath);

        // Load the created DOCX file.
        Document loadedDoc = new Document(inputPath);

        // Create a minimal Hunspell hyphenation dictionary for English (US).
        const string dictPath = "hyph_en_US.dic";
        File.WriteAllText(dictPath,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Register the dictionary.
        Hyphenation.RegisterDictionary("en-US", dictPath);

        // Enable automatic hyphenation.
        loadedDoc.HyphenationOptions.AutoHyphenation = true;

        // Save the hyphenated document to PDF.
        const string outputPath = "hyphenated.pdf";
        loadedDoc.Save(outputPath);

        // Validate that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The hyphenated PDF was not created.");
    }
}
