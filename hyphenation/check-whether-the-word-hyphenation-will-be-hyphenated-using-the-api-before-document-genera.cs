using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a minimal hyphenation dictionary for English (US) that contains the word "hyphenation".
        const string dictFileName = "hyph_en_US.dic";
        File.WriteAllText(dictFileName,
            "UTF-8\n" +
            "hyphenation=hy-phen-a-tion\n");

        // Register the dictionary for the "en-US" locale.
        Hyphenation.RegisterDictionary("en-US", dictFileName);

        // Verify that the dictionary is registered.
        bool isRegistered = Hyphenation.IsDictionaryRegistered("en-US");
        if (!isRegistered)
            throw new InvalidOperationException("Hyphenation dictionary was not registered.");

        // Create a document with narrow page width to force line wrapping.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("The word hyphenation may be split across lines when hyphenation is enabled.");

        // Narrow the page to make hyphenation visible.
        doc.FirstSection.PageSetup.PageWidth = 200;
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Save the document as PDF.
        const string outputFile = "HyphenationCheck.pdf";
        doc.Save(outputFile, SaveFormat.Pdf);

        // Validate that the output file was created.
        if (!File.Exists(outputFile))
            throw new InvalidOperationException("Expected output file was not created.");
    }
}
