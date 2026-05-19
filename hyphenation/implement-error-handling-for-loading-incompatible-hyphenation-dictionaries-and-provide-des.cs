using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using static Aspose.Words.Hyphenation; // Hyphenation is a static class, use static import

public class Program
{
    public static void Main()
    {
        // File names used in the example
        const string validDictPath = "valid_hyph_en_US.dic";
        const string invalidDictPath = "invalid_hyph_en_US.dic";
        const string outputPdf = "hyphenated.pdf";

        // Clean any previous artifacts
        if (File.Exists(validDictPath)) File.Delete(validDictPath);
        if (File.Exists(invalidDictPath)) File.Delete(invalidDictPath);
        if (File.Exists(outputPdf)) File.Delete(outputPdf);

        // -----------------------------------------------------------------
        // 1. Create a minimal valid hyphenation dictionary.
        //    The first line must be the encoding, followed by word=pattern lines.
        // -----------------------------------------------------------------
        File.WriteAllText(validDictPath,
            "UTF-8\nextraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n");

        // Register the valid dictionary – this should succeed.
        try
        {
            RegisterDictionary("en-US", validDictPath);
            Console.WriteLine("Valid dictionary registered successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to register valid dictionary: {ex.Message}");
        }

        // Verify registration state
        Console.WriteLine($"Is valid dictionary registered? {IsDictionaryRegistered("en-US")}");

        // -----------------------------------------------------------------
        // 2. Create an intentionally malformed dictionary to trigger an error.
        // -----------------------------------------------------------------
        File.WriteAllText(invalidDictPath,
            "This is not a valid hyphenation dictionary content");

        // Attempt to register the malformed dictionary and handle the exception.
        try
        {
            RegisterDictionary("en-US", invalidDictPath);
            Console.WriteLine("Invalid dictionary registered (unexpected).");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading incompatible dictionary: {ex.Message}");
        }

        // -----------------------------------------------------------------
        // 3. Build a document that will use hyphenation.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a font size that forces line wrapping and set the locale to match the dictionary.
        builder.Font.Size = 24;
        builder.Font.LocaleId = new CultureInfo("en-US").LCID;

        // Write a paragraph containing words that have hyphenation patterns.
        builder.Writeln("extraordinarycharacteristically internationalization communication");

        // Enable automatic hyphenation for the document.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Narrow the page width to make hyphenation visible.
        doc.FirstSection.PageSetup.PageWidth = 200;
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Save the document as PDF.
        doc.Save(outputPdf, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (File.Exists(outputPdf))
            Console.WriteLine($"Document saved to '{outputPdf}'.");
        else
            throw new InvalidOperationException("Expected PDF output file was not created.");
    }
}
