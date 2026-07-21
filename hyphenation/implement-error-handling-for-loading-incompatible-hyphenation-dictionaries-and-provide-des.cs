using System;
using System.Globalization;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // File names used in the example.
        const string validDictPath = "valid_hyph_en_US.dic";
        const string invalidDictPath = "invalid_hyph_en_US.dic";
        const string outputPdf = "Hyphenated.pdf";

        // -----------------------------------------------------------------
        // Create a minimal valid hyphenation dictionary.
        // The first line defines the encoding, subsequent lines define word‑hyphenation pairs.
        // -----------------------------------------------------------------
        File.WriteAllText(validDictPath,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Create an invalid (empty) dictionary to provoke an error.
        File.WriteAllText(invalidDictPath, string.Empty);

        // -----------------------------------------------------------------
        // Attempt to register the invalid dictionary.
        // The RegisterDictionary method throws if the file cannot be parsed.
        // -----------------------------------------------------------------
        try
        {
            Hyphenation.RegisterDictionary("en-US", invalidDictPath);
            Console.WriteLine("Invalid dictionary registered unexpectedly.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to register invalid dictionary: {ex.Message}");
        }

        // -----------------------------------------------------------------
        // Register the valid dictionary and verify that it is indeed registered.
        // -----------------------------------------------------------------
        try
        {
            Hyphenation.RegisterDictionary("en-US", validDictPath);
            Console.WriteLine("Valid dictionary registered successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to register valid dictionary: {ex.Message}");
            // If registration fails, further processing would be meaningless.
            return;
        }

        // Double‑check registration status.
        if (!Hyphenation.IsDictionaryRegistered("en-US"))
        {
            Console.WriteLine("Dictionary registration check failed.");
            return;
        }

        // -----------------------------------------------------------------
        // Build a document that forces hyphenation by using a narrow page width.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow page layout.
        doc.FirstSection.PageSetup.PageWidth = 200; // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.HyphenateCaps = true;

        // Write a line containing words that exist in the dictionary.
        builder.Font.Size = 24;
        builder.Writeln("extraordinarycharacteristically internationalization communication");

        // -----------------------------------------------------------------
        // Save the document as PDF.
        // -----------------------------------------------------------------
        doc.Save(outputPdf, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(outputPdf))
            throw new InvalidOperationException("Expected PDF output file was not created.");

        Console.WriteLine($"Document saved successfully to '{outputPdf}'.");
    }
}
