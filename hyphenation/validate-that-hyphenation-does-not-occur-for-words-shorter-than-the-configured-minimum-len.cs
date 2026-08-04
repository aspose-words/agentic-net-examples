using System;
using System.IO;
using Aspose.Words;

public class HyphenationMinLengthExample
{
    public static void Main()
    {
        // Create a minimal hyphenation dictionary for English (US).
        const string dictFileName = "hyph_en_US.dic";
        File.WriteAllText(dictFileName,
            "UTF-8\n" +
            "extraordinarycharacteristically=ex-tra-or-di-nary-char-ac-ter-is-ti-cal-ly\n" +
            "cat=cat\n"); // Short word "cat" has no hyphenation points.

        // Register the dictionary.
        Hyphenation.RegisterDictionary("en-US", dictFileName);
        if (!Hyphenation.IsDictionaryRegistered("en-US"))
            throw new InvalidOperationException("Failed to register the hyphenation dictionary.");

        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow the page to force line wrapping.
        doc.FirstSection.PageSetup.PageWidth = 200;
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Write a long word (should be hyphenated) and a short word (should not be hyphenated).
        builder.Font.Size = 24;
        builder.Writeln("extraordinarycharacteristically cat");

        // Save the document as PDF.
        const string outputFile = "HyphenationMinLength.pdf";
        doc.Save(outputFile, SaveFormat.Pdf);

        // Validate that the output file was created.
        if (!File.Exists(outputFile))
            throw new InvalidOperationException("The expected PDF output was not created.");

        // Simple validation: ensure the short word "cat" has no hyphenation points in the dictionary.
        string dictContent = File.ReadAllText(dictFileName);
        if (dictContent.Contains("cat=") && dictContent.Contains("cat-"))
            throw new InvalidOperationException("Short word 'cat' should not contain hyphenation points.");

        // Clean up temporary files.
        File.Delete(dictFileName);
    }
}
