using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Paths for artifacts
        const string dictionaryPath = "hyph_en_US.dic";
        const string firstPdfPath = "hyphenated_v1.pdf";
        const string secondPdfPath = "hyphenated_v2.pdf";

        // Step 1: Create initial hyphenation dictionary
        string initialDictionaryContent =
            "UTF-8\n" +
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "internationalization=in-ter-na-tion-al-i-za-tion\n" +
            "communication=com-mu-ni-ca-tion\n";
        File.WriteAllText(dictionaryPath, initialDictionaryContent);

        // Step 2: Register the dictionary for en-US locale
        Aspose.Words.Hyphenation.RegisterDictionary("en-US", dictionaryPath);

        // Step 3: Build a document that will trigger hyphenation
        Document docV1 = new Document();
        DocumentBuilder builderV1 = new DocumentBuilder(docV1);
        // Narrow page width to force line breaks
        docV1.FirstSection.PageSetup.PageWidth = 200;
        docV1.FirstSection.PageSetup.LeftMargin = 20;
        docV1.FirstSection.PageSetup.RightMargin = 20;
        // Sample text containing words from the dictionary
        builderV1.Writeln("extraordinarycharacteristically internationalization communication");

        // Step 4: Save the first PDF
        docV1.Save(firstPdfPath, SaveFormat.Pdf);
        if (!File.Exists(firstPdfPath))
            throw new InvalidOperationException($"Failed to create {firstPdfPath}");

        // Step 5: Simulate CI pipeline dictionary update
        string updatedDictionaryContent =
            initialDictionaryContent +
            "hyperresponsibility=hy-per-re-spon-si-bil-i-ty\n";
        File.WriteAllText(dictionaryPath, updatedDictionaryContent);

        // Re‑register the updated dictionary
        Aspose.Words.Hyphenation.RegisterDictionary("en-US", dictionaryPath);

        // Step 6: Build a second document that uses the new word
        Document docV2 = new Document();
        DocumentBuilder builderV2 = new DocumentBuilder(docV2);
        docV2.FirstSection.PageSetup.PageWidth = 200;
        docV2.FirstSection.PageSetup.LeftMargin = 20;
        docV2.FirstSection.PageSetup.RightMargin = 20;
        builderV2.Writeln("hyperresponsibility extraordinarycharacteristically");

        // Step 7: Save the second PDF
        docV2.Save(secondPdfPath, SaveFormat.Pdf);
        if (!File.Exists(secondPdfPath))
            throw new InvalidOperationException($"Failed to create {secondPdfPath}");

        // Cleanup (optional)
        // File.Delete(dictionaryPath);
    }
}
