using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class HyphenationMinimumLengthDemo
{
    public static void Main()
    {
        // Create a simple document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add a paragraph with a long word (eligible for hyphenation) and a short word (should not hyphenate).
        builder.Writeln("extraordinarycharacteristically cat");

        // Narrow the page width to force line wrapping and potential hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 200; // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Create a minimal hyphenation dictionary for English (US).
        const string dictFileName = "hyph_en_US.dic";
        File.WriteAllText(dictFileName,
            "UTF-8\n" +
            "extraordinarycharacteristically=ex-tra-or-di-nary-char-ac-ter-is-ti-cal-ly\n");

        // Register the dictionary.
        Hyphenation.RegisterDictionary("en-US", dictFileName);

        // Verify that the dictionary is registered.
        if (!Hyphenation.IsDictionaryRegistered("en-US"))
            throw new InvalidOperationException("Hyphenation dictionary was not registered.");

        // Save the document to PDF.
        const string outputPdf = "HyphenationDemo.pdf";
        doc.Save(outputPdf, SaveFormat.Pdf);

        // Validate that the PDF was created.
        if (!File.Exists(outputPdf))
            throw new InvalidOperationException("Expected PDF output was not created.");

        // Clean up temporary files (optional).
        // File.Delete(dictFileName);
        // File.Delete(outputPdf);
    }
}
