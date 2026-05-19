using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a minimal hyphenation dictionary for English (US) that includes the word "hyphenation".
        const string dictFileName = "hyph_en_US.dic";
        const string dictContent = "UTF-8\nhyphenation=hy-phen-a-tion\n";
        File.WriteAllText(dictFileName, dictContent);

        // Register the dictionary for the "en-US" language.
        // The Hyphenation class resides directly in the Aspose.Words namespace.
        Aspose.Words.Hyphenation.RegisterDictionary("en-US", dictFileName);

        // Verify that the dictionary is registered.
        if (!Aspose.Words.Hyphenation.IsDictionaryRegistered("en-US"))
            throw new InvalidOperationException("Hyphenation dictionary for en-US was not registered.");

        // Create a document containing the word "hyphenation" repeated to force line wrapping.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Size = 24;
        builder.Writeln("hyphenation hyphenation hyphenation hyphenation hyphenation hyphenation hyphenation hyphenation");

        // Narrow the page width to increase the chance of hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 200;
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Save the document as PDF.
        const string outputFile = "hyphenated.pdf";
        doc.Save(outputFile);

        // Validate that the PDF was created.
        if (!File.Exists(outputFile))
            throw new InvalidOperationException("Expected output PDF was not created.");

        // Output the result of the hyphenation check.
        Console.WriteLine($"Hyphenation dictionary registered: {Aspose.Words.Hyphenation.IsDictionaryRegistered("en-US")}");
        Console.WriteLine($"Document saved to: {Path.GetFullPath(outputFile)}");
    }
}
