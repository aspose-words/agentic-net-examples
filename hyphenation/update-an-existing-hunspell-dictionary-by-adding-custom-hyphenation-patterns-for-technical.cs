using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a local Hunspell hyphenation dictionary file.
        string dictPath = Path.Combine(Directory.GetCurrentDirectory(), "hyph_en_US.dic");
        string dictionaryContent =
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n" +
            "microprocessor=mi-cro-pro-ces-sor\n" +
            "quantumcomputing=quan-tum-com-put-ing\n";

        File.WriteAllText(dictPath, dictionaryContent);

        // Register the dictionary for the en-US locale.
        Aspose.Words.Hyphenation.RegisterDictionary("en-US", dictPath);

        // Build a document with narrow page width to force hyphenation.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        doc.FirstSection.PageSetup.PageWidth = 200; // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Sample text containing the technical terms.
        builder.Writeln(
            "The microprocessor architecture and quantumcomputing algorithms are essential for modern technology. " +
            "Extraordinarycharacteristically, internationalization, and communication also play vital roles.");

        // Save the result as PDF.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "hyphenated_output.pdf");
        doc.Save(outputPath, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Expected output PDF was not created.");
    }
}
