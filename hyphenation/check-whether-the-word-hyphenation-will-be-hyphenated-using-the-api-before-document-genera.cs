using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a minimal hyphenation dictionary for English (US) that includes the word "hyphenation".
        const string dictionaryPath = "hyph_en_US.dic";
        File.WriteAllText(dictionaryPath, "UTF-8\nhyphenation=hy-phen-ation\n");

        // Register the dictionary with Aspose.Words.
        Hyphenation.RegisterDictionary("en-US", dictionaryPath);

        // Verify that the dictionary was successfully registered.
        if (!Hyphenation.IsDictionaryRegistered("en-US"))
            throw new InvalidOperationException("Failed to register the hyphenation dictionary.");

        // Create a new document and enable automatic hyphenation.
        Document doc = new Document();
        doc.HyphenationOptions.AutoHyphenation = true;

        // Configure the page layout to force line wrapping, making hyphenation visible.
        doc.FirstSection.PageSetup.PageWidth = 200;
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Add the word "hyphenation" to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Size = 24;
        builder.Writeln("hyphenation");

        // Save the document to trigger layout processing.
        const string outputPath = "HyphenationCheck.pdf";
        doc.Save(outputPath);

        // Ensure the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output PDF was not created.");

        // Indicate successful execution.
        Console.WriteLine("Hyphenation dictionary registered and document generated successfully.");
    }
}
