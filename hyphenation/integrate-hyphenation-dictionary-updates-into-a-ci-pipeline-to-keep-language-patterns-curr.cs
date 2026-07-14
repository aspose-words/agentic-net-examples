using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Paths for artifacts
        const string dictPath = "hyph_en_US.dic";
        const string pdfBeforePath = "hyphenated_before.pdf";
        const string pdfAfterPath = "hyphenated_after.pdf";

        // Ensure a clean environment
        if (File.Exists(dictPath)) File.Delete(dictPath);
        if (File.Exists(pdfBeforePath)) File.Delete(pdfBeforePath);
        if (File.Exists(pdfAfterPath)) File.Delete(pdfAfterPath);

        // -----------------------------------------------------------------
        // Step 1: Create an initial hyphenation dictionary file.
        // The dictionary follows the OpenOffice format: first line is the encoding,
        // subsequent lines contain word=hyphenation-points.
        // -----------------------------------------------------------------
        File.WriteAllText(dictPath,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n");

        // Register the dictionary for the "en-US" locale.
        Hyphenation.RegisterDictionary("en-US", dictPath);

        // -----------------------------------------------------------------
        // Step 2: Build a sample document that will trigger hyphenation.
        // -----------------------------------------------------------------
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Use a narrow page width to force line wrapping.
        doc.FirstSection.PageSetup.PageWidth = 200; // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenationZone = 360; // 0.25 inch
        doc.HyphenationOptions.HyphenateCaps = true;

        // Set the paragraph locale to match the dictionary language.
        builder.Font.LocaleId = new CultureInfo("en-US").LCID;
        builder.Font.Size = 12;
        builder.Writeln("extraordinarycharacteristically internationalization communication");

        // Save the first PDF using the initial dictionary.
        doc.Save(pdfBeforePath, SaveFormat.Pdf);
        if (!File.Exists(pdfBeforePath))
            throw new InvalidOperationException($"Failed to create '{pdfBeforePath}'.");

        // -----------------------------------------------------------------
        // Step 3: Simulate a CI pipeline update – modify the dictionary.
        // -----------------------------------------------------------------
        // Append a new hyphenation pattern for the word "communication".
        File.AppendAllText(dictPath,
            "communication=com-mu-ni-ca-tion\n");

        // Re‑register the updated dictionary.
        Hyphenation.UnregisterDictionary("en-US");
        Hyphenation.RegisterDictionary("en-US", dictPath);

        // Re‑save the document to a new PDF to reflect the updated patterns.
        doc.Save(pdfAfterPath, SaveFormat.Pdf);
        if (!File.Exists(pdfAfterPath))
            throw new InvalidOperationException($"Failed to create '{pdfAfterPath}'.");

        // -----------------------------------------------------------------
        // Step 4: Validation – both PDFs should exist.
        // -----------------------------------------------------------------
        Console.WriteLine($"Generated PDFs:\n  {pdfBeforePath}\n  {pdfAfterPath}");
    }
}
