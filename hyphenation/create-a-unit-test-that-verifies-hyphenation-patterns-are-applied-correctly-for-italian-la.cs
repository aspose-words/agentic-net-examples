using System;
using System.Globalization;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Path for the Italian hyphenation dictionary.
        const string dictionaryPath = "hyph_it_IT.dic";

        // Create a minimal valid dictionary file.
        // The format is: first line "UTF-8", then each entry "word=pattern".
        File.WriteAllText(dictionaryPath, "UTF-8\ncasa=ca-sa\n");

        // Register the dictionary for the Italian locale.
        Hyphenation.RegisterDictionary("it-IT", dictionaryPath);

        // Verify that the dictionary was registered successfully.
        if (!Hyphenation.IsDictionaryRegistered("it-IT"))
            throw new InvalidOperationException("Italian hyphenation dictionary was not registered.");

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Configure a narrow page width to force line wrapping.
        doc.FirstSection.PageSetup.PageWidth = 200;   // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Enable automatic hyphenation for the document.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Set the font locale to Italian.
        builder.Font.LocaleId = new CultureInfo("it-IT").LCID;
        builder.Font.Size = 12;

        // Write a paragraph containing many occurrences of the word "casa".
        // The dictionary defines a hyphenation point "ca-sa" for this word.
        string repeatedWord = string.Join(" ", Enumerable.Repeat("casa", 50));
        builder.Writeln(repeatedWord);

        // Save the document to PDF.
        const string outputPath = "hyphenated_it.pdf";
        doc.Save(outputPath, SaveFormat.Pdf);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The hyphenated PDF was not created.");

        // If we reach this point, the test succeeded.
        Console.WriteLine("Hyphenation test completed successfully.");
    }
}
