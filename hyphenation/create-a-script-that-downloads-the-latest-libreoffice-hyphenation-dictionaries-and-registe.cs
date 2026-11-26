using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a minimal hyphenation dictionary file locally.
        const string dictionaryPath = "hyph_en_US.dic";
        string dictionaryContent =
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n";

        File.WriteAllText(dictionaryPath, dictionaryContent);

        // Register the dictionary for the en-US locale.
        Hyphenation.RegisterDictionary("en-US", dictionaryPath);

        // Build a document that contains words from the dictionary.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow page width forces line wrapping and hyphenation.
        Section section = doc.FirstSection;
        PageSetup setup = section.PageSetup;
        setup.PageWidth = 200; // points (~2.78 inches)
        setup.LeftMargin = 20;
        setup.RightMargin = 20;

        // Write sample text that can be hyphenated.
        builder.Writeln("extraordinarycharacteristically internationalization communication");

        // Save the document as PDF.
        const string outputPath = "hyphenated.pdf";
        doc.Save(outputPath, SaveFormat.Pdf);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Expected output file '{outputPath}' was not created.");
        }
    }
}
