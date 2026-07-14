using System;
using System.Globalization;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define file names.
        const string dictionaryFile = "hyph_ar_SA.dic";
        const string outputFile = "HyphenatedArabic.pdf";

        // Create a minimal Arabic hyphenation dictionary.
        // Format: first line is the encoding, subsequent lines are word=hyphenated-parts.
        // The example word "البرمجة" (programming) is split as "ال-برم-جة".
        File.WriteAllText(dictionaryFile,
            "UTF-8\n" +
            "البرمجة=ال-برم-جة\n" +
            "المعلومات=الم-علو-مات\n");

        // Register the Arabic dictionary with the appropriate language code.
        Aspose.Words.Hyphenation.RegisterDictionary("ar-SA", dictionaryFile);

        // Verify that the dictionary was registered.
        if (!Aspose.Words.Hyphenation.IsDictionaryRegistered("ar-SA"))
            throw new InvalidOperationException("Failed to register the Arabic hyphenation dictionary.");

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the document locale to Arabic (Saudi Arabia) and enable right‑to‑left layout.
        builder.Font.LocaleId = new CultureInfo("ar-SA").LCID;
        builder.ParagraphFormat.Bidi = true; // Right‑to‑left paragraph.

        // Write a long Arabic sentence that will require line wrapping.
        // The sentence repeats words that are present in the dictionary to trigger hyphenation.
        string arabicText = "البرمجة هي عملية كتابة التعليمات البرمجية. " +
                            "المعلومات هي أساس أي نظام. " +
                            "البرمجة والبرمجة والبرمجة والبرمجة والالبرمجة والالبرمجة والالبرمجة.";
        builder.Writeln(arabicText);

        // Enable automatic hyphenation for the document.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Reduce the page width to force line breaks and make hyphenation visible.
        doc.FirstSection.PageSetup.PageWidth = 300; // Points (approx 4.2 cm)
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Save the document to PDF.
        doc.Save(outputFile, SaveFormat.Pdf);

        // Validate that the output file was created.
        if (!File.Exists(outputFile))
            throw new InvalidOperationException("The expected PDF output was not created.");

        // Optional cleanup of the temporary dictionary file.
        // File.Delete(dictionaryFile);
    }
}
