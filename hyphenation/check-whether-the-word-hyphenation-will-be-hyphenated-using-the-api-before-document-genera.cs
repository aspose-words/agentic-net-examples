using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class HyphenationCheckExample
{
    public static void Main()
    {
        // Prepare a minimal hyphenation dictionary for English (US) that contains the word "hyphenation".
        const string dictFileName = "hyph_en_US.dic";
        const string dictContent = "UTF-8\nhyphenation=hy-phen-a-tion\n";
        File.WriteAllText(dictFileName, dictContent);

        // Register the dictionary for the "en-US" language.
        Hyphenation.RegisterDictionary("en-US", dictFileName);

        // Verify that the dictionary is registered.
        if (!Hyphenation.IsDictionaryRegistered("en-US"))
            throw new InvalidOperationException("Hyphenation dictionary for en-US was not registered.");

        // At this point we can assume that the word "hyphenation" can be hyphenated
        // because it is present in the registered dictionary.
        Console.WriteLine("Hyphenation dictionary registered. The word 'hyphenation' is hyphenatable.");

        // Create a document that will use automatic hyphenation.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;
        // Set a narrow page width to force line breaks.
        doc.FirstSection.PageSetup.PageWidth = 200;
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Write a paragraph containing the word "hyphenation" multiple times.
        builder.Font.Size = 12;
        builder.Writeln("hyphenation hyphenation hyphenation hyphenation hyphenation hyphenation hyphenation hyphenation hyphenation hyphenation");

        // Save the document to PDF to visualize hyphenation (optional).
        const string outputFile = "HyphenatedOutput.pdf";
        doc.Save(outputFile, SaveFormat.Pdf);

        // Validate that the output file was created.
        if (!File.Exists(outputFile))
            throw new InvalidOperationException("The PDF output file was not created.");

        Console.WriteLine($"Document saved to '{outputFile}'.");
    }
}
