using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Paths for the dictionary and output files.
        const string dictionaryPath = "hyph_en_US.dic";
        const string outputPathV1 = "hyphenated_v1.pdf";
        const string outputPathV2 = "hyphenated_v2.pdf";

        // -----------------------------------------------------------------
        // Step 1: Create an initial hyphenation dictionary file.
        // -----------------------------------------------------------------
        File.WriteAllText(dictionaryPath,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Register the dictionary for the "en-US" locale.
        Hyphenation.RegisterDictionary("en-US", dictionaryPath);

        // Verify registration.
        if (!Hyphenation.IsDictionaryRegistered("en-US"))
            throw new InvalidOperationException("Dictionary registration failed.");

        // -----------------------------------------------------------------
        // Step 2: Build a sample document that will trigger hyphenation.
        // -----------------------------------------------------------------
        Document docV1 = new Document();
        DocumentBuilder builder = new DocumentBuilder(docV1);

        // Narrow page width forces line wrapping.
        docV1.FirstSection.PageSetup.PageWidth = 200;
        docV1.FirstSection.PageSetup.LeftMargin = 20;
        docV1.FirstSection.PageSetup.RightMargin = 20;

        // Write a paragraph containing words defined in the dictionary.
        builder.Font.Size = 12;
        builder.Writeln(
            "extraordinarycharacteristically internationalization communication " +
            "demonstration of hyphenation handling in a CI pipeline scenario.");

        // Enable automatic hyphenation.
        docV1.HyphenationOptions.AutoHyphenation = true;

        // Save the first version of the PDF.
        docV1.Save(outputPathV1, SaveFormat.Pdf);

        // Validate output.
        if (!File.Exists(outputPathV1))
            throw new InvalidOperationException($"Expected file '{outputPathV1}' was not created.");

        // -----------------------------------------------------------------
        // Step 3: Simulate a CI pipeline update – modify the dictionary.
        // -----------------------------------------------------------------
        // Append a new hyphenation pattern to the dictionary file.
        File.AppendAllText(dictionaryPath,
            "demonstration=de-mon-stra-tion\n");

        // Unregister the old dictionary and register the updated one.
        Hyphenation.UnregisterDictionary("en-US");
        Hyphenation.RegisterDictionary("en-US", dictionaryPath);

        // Verify re‑registration.
        if (!Hyphenation.IsDictionaryRegistered("en-US"))
            throw new InvalidOperationException("Updated dictionary registration failed.");

        // -----------------------------------------------------------------
        // Step 4: Build a second document using the updated dictionary.
        // -----------------------------------------------------------------
        Document docV2 = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(docV2);

        // Apply the same page setup to keep layout consistent.
        docV2.FirstSection.PageSetup.PageWidth = 200;
        docV2.FirstSection.PageSetup.LeftMargin = 20;
        docV2.FirstSection.PageSetup.RightMargin = 20;

        builder2.Font.Size = 12;
        builder2.Writeln(
            "extraordinarycharacteristically internationalization communication " +
            "demonstration of hyphenation handling after dictionary update.");

        // Enable automatic hyphenation.
        docV2.HyphenationOptions.AutoHyphenation = true;

        // Save the second version of the PDF.
        docV2.Save(outputPathV2, SaveFormat.Pdf);

        // Validate output.
        if (!File.Exists(outputPathV2))
            throw new InvalidOperationException($"Expected file '{outputPathV2}' was not created.");

        // -----------------------------------------------------------------
        // Completion message (optional, not interactive).
        // -----------------------------------------------------------------
        Console.WriteLine("Hyphenation processing completed successfully.");
    }
}
