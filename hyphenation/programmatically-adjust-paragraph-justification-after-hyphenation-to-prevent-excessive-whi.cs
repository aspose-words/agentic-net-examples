using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a minimal hyphenation dictionary for English (en-US).
        const string dictFileName = "hyph_en_US.dic";
        string dictContent =
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n";

        File.WriteAllText(dictFileName, dictContent);

        // Register the dictionary with Aspose.Words.
        Aspose.Words.Hyphenation.RegisterDictionary("en-US", dictFileName);

        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow the page width to force line wrapping and hyphenation.
        Section section = doc.FirstSection;
        section.PageSetup.PageWidth = 200; // points (~2.78 inches)
        section.PageSetup.LeftMargin = 20;
        section.PageSetup.RightMargin = 20;

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Write a paragraph with long words that can be hyphenated.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Justify;
        builder.Writeln("extraordinarycharacteristically internationalization communication");

        // Save the document as PDF.
        const string outputFile = "hyphenated.pdf";
        doc.Save(outputFile, SaveFormat.Pdf);

        // Validate that the output file was created.
        if (!File.Exists(outputFile))
        {
            throw new InvalidOperationException($"Expected output file '{outputFile}' was not created.");
        }

        // Clean up temporary dictionary file (optional).
        if (File.Exists(dictFileName))
        {
            File.Delete(dictFileName);
        }
    }
}
