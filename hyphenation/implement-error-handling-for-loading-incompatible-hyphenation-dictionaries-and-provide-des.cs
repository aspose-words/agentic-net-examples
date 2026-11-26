using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a simple document with long words that can be hyphenated.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("extraordinarycharacteristically internationalization communication");

        // Narrow page width to force line wrapping and hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 200;
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Prepare a valid hyphenation dictionary.
        const string validDicPath = "hyph_en_US_valid.dic";
        File.WriteAllText(validDicPath,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Register the valid dictionary – this should succeed.
        Aspose.Words.Hyphenation.RegisterDictionary("en-US", validDicPath);

        // Save the document as PDF to demonstrate successful hyphenation.
        const string validPdfPath = "valid_hyphenated.pdf";
        doc.Save(validPdfPath, SaveFormat.Pdf);
        if (!File.Exists(validPdfPath))
            throw new InvalidOperationException("The PDF with valid hyphenation was not created.");

        // Prepare an invalid hyphenation dictionary (missing the required header line).
        const string invalidDicPath = "hyph_invalid.dic";
        File.WriteAllText(invalidDicPath,
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n");

        // Attempt to register the invalid dictionary and handle the error.
        try
        {
            Aspose.Words.Hyphenation.RegisterDictionary("en-US", invalidDicPath);
            // If no exception is thrown, the example should indicate unexpected success.
            Console.WriteLine("Warning: Invalid dictionary was registered without error.");
        }
        catch (Exception ex)
        {
            // Provide a descriptive message about the failure.
            Console.WriteLine($"Failed to register hyphenation dictionary '{invalidDicPath}': {ex.Message}");
        }

        // Optional clean‑up of temporary files.
        // File.Delete(validDicPath);
        // File.Delete(invalidDicPath);
        // File.Delete(validPdfPath);
    }
}
