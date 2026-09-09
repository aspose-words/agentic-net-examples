using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Create a minimal hyphenation dictionary for English (US).
        const string dictFileName = "hyph_en_US.dic";
        File.WriteAllText(dictFileName,
            "UTF-8\n" +
            "shortword=short-word\n"); // Defines a hyphenation point for a short word.

        // Register the dictionary for the "en-US" locale.
        Hyphenation.RegisterDictionary("en-US", dictFileName);

        // Verify that the dictionary was registered successfully.
        if (!Hyphenation.IsDictionaryRegistered("en-US"))
            throw new InvalidOperationException("Hyphenation dictionary was not registered.");

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow the page width to force line wrapping.
        doc.FirstSection.PageSetup.PageWidth = 200;
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Write a paragraph containing the short word multiple times.
        // The word "shortword" is 9 characters long; Aspose.Words does not hyphenate
        // words shorter than the dictionary's defined hyphenation point by default.
        builder.Font.Size = 24;
        builder.Writeln("shortword shortword shortword shortword shortword shortword shortword shortword");

        // Save the document to PDF.
        const string outputFile = "Hyphenation_MinLength.pdf";
        doc.Save(outputFile, SaveFormat.Pdf);

        // Verify that the PDF file was created.
        if (!File.Exists(outputFile))
            throw new InvalidOperationException("The PDF output file was not created.");

        // Clean up the temporary dictionary file.
        File.Delete(dictFileName);
    }
}
