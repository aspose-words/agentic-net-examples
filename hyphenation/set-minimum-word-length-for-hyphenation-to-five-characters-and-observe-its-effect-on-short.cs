using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class HyphenationMinWordLengthExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set a narrow page width to force line wrapping and possible hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 300; // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Register a minimal English hyphenation dictionary.
        const string dictFileName = "hyph_en_US.dic";
        // The dictionary format: first line is "UTF-8", subsequent lines are word=hyphenation-points.
        // Include a long word that can be hyphenated.
        File.WriteAllText(dictFileName,
            "UTF-8\nextraordinarycharacteristically=ex-tra-or-di-nary-char-ac-ter-is-ti-cal-ly\ncommunication=com-mu-ni-ca-tion\n");

        // Register the dictionary for the "en-US" locale.
        Hyphenation.RegisterDictionary("en-US", dictFileName);

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;
        // Set the minimum word length for hyphenation to 5 characters.
        // Aspose.Words automatically respects a minimum length; we simulate it by ensuring the dictionary
        // contains entries only for words longer than 5 characters.
        // No explicit property exists, so this setting is implicit.

        // Write a paragraph containing short and long words.
        // Short words (<=4 characters) should not be hyphenated.
        // Long words (>5 characters) that exist in the dictionary may be hyphenated.
        builder.Font.Size = 12;
        builder.Writeln("Short words: cat dog sun. Long word: extraordinarycharacteristically communication.");

        // Save the document to PDF to visualize hyphenation.
        const string outputFile = "HyphenationMinWordLength.pdf";
        doc.Save(outputFile, SaveFormat.Pdf);

        // Validate that the output file was created.
        if (!File.Exists(outputFile))
            throw new InvalidOperationException("The PDF output file was not created.");

        // Clean up temporary dictionary file.
        File.Delete(dictFileName);
    }
}
