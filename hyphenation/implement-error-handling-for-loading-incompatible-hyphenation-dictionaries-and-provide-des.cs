using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading; // For WarningInfoCollection

public class Program
{
    public static void Main()
    {
        // Create a simple document with long words that can be hyphenated.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Size = 24;
        builder.Writeln("extraordinarycharacteristically internationalization communication");

        // Narrow the page width so the words wrap and hyphenation can be observed.
        doc.FirstSection.PageSetup.PageWidth = 200;
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Create a deliberately malformed hyphenation dictionary.
        const string dictPath = "invalid_hyph_en_US.dic";
        File.WriteAllText(dictPath,
            // Missing the required header line ("UTF-8") and contains garbage.
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "this line is not a valid pattern");

        // Set up a warning collector to capture any warnings during registration.
        WarningInfoCollection warnings = new WarningInfoCollection();
        Hyphenation.WarningCallback = warnings;

        // Attempt to register the malformed dictionary and handle errors gracefully.
        try
        {
            Hyphenation.RegisterDictionary("en-US", dictPath);
            Console.WriteLine("Dictionary registered successfully.");
        }
        catch (Exception ex)
        {
            // Provide a clear, descriptive message for the failure.
            Console.WriteLine($"Failed to register hyphenation dictionary for 'en-US': {ex.Message}");
        }

        // Report any warnings that were raised during the registration attempt.
        if (warnings.Count > 0)
        {
            Console.WriteLine("Hyphenation warnings:");
            foreach (WarningInfo warning in warnings)
            {
                Console.WriteLine($"- {warning.WarningType}: {warning.Description}");
            }
        }

        // Save the document. If the dictionary was invalid, hyphenation will not be applied,
        // but the document will still be saved.
        const string outputPath = "HyphenatedOutput.pdf";
        doc.Save(outputPath);

        // Verify that the output file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved to '{outputPath}'.");
        }
        else
        {
            throw new InvalidOperationException("The expected output PDF was not created.");
        }
    }
}
