using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Path for the temporary hyphenation dictionary.
        const string dictPath = "hyph_it_IT.dic";

        // Create a minimal Italian hyphenation dictionary.
        // The first line must be the encoding, followed by word=hyphenated‑pattern lines.
        File.WriteAllText(dictPath,
            "UTF-8\nextraordinariamente=ex-tra-or-di-nar-io-men-te\n");

        // Register the dictionary for the Italian locale (it-IT).
        Aspose.Words.Hyphenation.RegisterDictionary("it-IT", dictPath);

        // Create a new document and a builder to add content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;
        // Optional: define the hyphenation zone (distance from the right margin where hyphenation may occur).
        doc.HyphenationOptions.HyphenationZone = 5;

        // Set the locale of the text to Italian (LocaleId 1040).
        builder.Font.LocaleId = 1040;

        // Narrow the page width to force line wrapping and thus hyphenation.
        Section section = doc.FirstSection;
        section.PageSetup.PageWidth = 200;   // points
        section.PageSetup.LeftMargin = 20;
        section.PageSetup.RightMargin = 20;

        // Write a long paragraph containing the word defined in the dictionary.
        const string word = "extraordinariamente";
        for (int i = 0; i < 30; i++)
        {
            builder.Writeln(word);
        }

        // Save the document as PDF.
        const string outputPath = "hyphenated_it.pdf";
        doc.Save(outputPath, SaveFormat.Pdf);

        // ----- Validation -----
        // Ensure the dictionary file was created.
        if (!File.Exists(dictPath))
            throw new InvalidOperationException("Hyphenation dictionary file was not created.");

        // Ensure the PDF output was created and is not empty.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Expected PDF output was not created.");

        if (new FileInfo(outputPath).Length == 0)
            throw new InvalidOperationException("Output PDF file is empty.");

        // Verify that automatic hyphenation is enabled.
        if (!doc.HyphenationOptions.AutoHyphenation)
            throw new InvalidOperationException("Automatic hyphenation is not enabled.");

        // Indicate successful completion.
        Console.WriteLine("Hyphenation test completed successfully.");
    }
}
