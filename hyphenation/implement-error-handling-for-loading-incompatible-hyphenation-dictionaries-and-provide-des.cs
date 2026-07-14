using System;
using System.IO;
using Aspose.Words;

public class HyphenationErrorHandlingExample
{
    public static void Main()
    {
        // File names for the dictionaries and the output PDF.
        const string validDictionaryPath = "hyph_en_US_valid.dic";
        const string invalidDictionaryPath = "hyph_en_US_invalid.dic";
        const string outputPdfPath = "hyphenated.pdf";

        // Create a minimal valid hyphenation dictionary.
        // The first line must be the encoding identifier (e.g., "UTF-8").
        // Subsequent lines contain word=hyphenation-points.
        File.WriteAllText(validDictionaryPath,
            "UTF-8\nextraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\ncommunication=com-mu-ni-ca-tion");

        // Create an invalid dictionary (missing the required header line).
        File.WriteAllText(invalidDictionaryPath,
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\ncommunication=com-mu-ni-ca-tion");

        // Attempt to register the valid dictionary.
        try
        {
            Hyphenation.RegisterDictionary("en-US", validDictionaryPath);
            Console.WriteLine($"Successfully registered dictionary: {validDictionaryPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error registering valid dictionary: {ex.Message}");
            // Abort if the valid dictionary cannot be loaded – further steps depend on it.
            return;
        }

        // Attempt to register the invalid dictionary and handle the expected failure.
        try
        {
            Hyphenation.RegisterDictionary("en-US", invalidDictionaryPath);
            Console.WriteLine($"Unexpectedly succeeded in registering invalid dictionary: {invalidDictionaryPath}");
        }
        catch (Exception ex)
        {
            // Provide a descriptive message for the failure.
            Console.WriteLine($"Failed to register dictionary '{invalidDictionaryPath}': {ex.Message}");
        }

        // Build a document that will demonstrate hyphenation.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a narrow page width to force line wrapping.
        doc.FirstSection.PageSetup.PageWidth = 200;
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Write a paragraph containing words that can be hyphenated.
        builder.Font.Size = 24;
        builder.Writeln(
            "extraordinarycharacteristically communication " +
            "extraordinarycharacteristically communication");

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Save the document to PDF.
        try
        {
            doc.Save(outputPdfPath);
            Console.WriteLine($"Document saved successfully: {outputPdfPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving document: {ex.Message}");
            return;
        }

        // Verify that the output file was created.
        if (!File.Exists(outputPdfPath))
            throw new InvalidOperationException($"Expected output file '{outputPdfPath}' was not created.");
    }
}
