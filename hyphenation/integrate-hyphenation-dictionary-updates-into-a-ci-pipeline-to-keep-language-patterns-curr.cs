using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Paths for the dictionary and output PDFs.
        const string dictionaryPath = "hyph_en_US.dic";
        const string pdfV1Path = "hyphenated_v1.pdf";
        const string pdfV2Path = "hyphenated_v2.pdf";

        // -----------------------------------------------------------------
        // Step 1: Create an initial hyphenation dictionary (simulating the
        //         state of the dictionary before a CI pipeline run).
        // -----------------------------------------------------------------
        File.WriteAllText(dictionaryPath,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n");

        // Register the dictionary for the "en-US" locale.
        Hyphenation.RegisterDictionary("en-US", dictionaryPath);

        // -----------------------------------------------------------------
        // Step 2: Build a sample document that will be hyphenated using the
        //         dictionary registered above.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow page width forces line wrapping and hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 200;
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Sample text containing a word that has a hyphenation pattern.
        builder.Writeln("extraordinarycharacteristically internationalization communication");

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Save the first PDF (baseline version).
        doc.Save(pdfV1Path);
        if (!File.Exists(pdfV1Path))
            throw new InvalidOperationException($"Failed to create {pdfV1Path}");

        // -----------------------------------------------------------------
        // Step 3: Simulate a CI pipeline update – modify the dictionary.
        // -----------------------------------------------------------------
        File.WriteAllText(dictionaryPath,
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "communication=com-mu-ni-ca-tion\n");

        // Unregister the old dictionary and register the updated one.
        Hyphenation.UnregisterDictionary("en-US");
        Hyphenation.RegisterDictionary("en-US", dictionaryPath);

        // -----------------------------------------------------------------
        // Step 4: Re‑save the document after the dictionary update.
        // -----------------------------------------------------------------
        // Force a layout rebuild so the new hyphenation rules are applied.
        doc.UpdatePageLayout();
        doc.Save(pdfV2Path);
        if (!File.Exists(pdfV2Path))
            throw new InvalidOperationException($"Failed to create {pdfV2Path}");

        // -----------------------------------------------------------------
        // Step 5: Simple validation – ensure both PDFs were produced.
        // -----------------------------------------------------------------
        Console.WriteLine($"Generated PDFs:\n  {pdfV1Path} ({new FileInfo(pdfV1Path).Length} bytes)\n  {pdfV2Path} ({new FileInfo(pdfV2Path).Length} bytes)");
    }
}
