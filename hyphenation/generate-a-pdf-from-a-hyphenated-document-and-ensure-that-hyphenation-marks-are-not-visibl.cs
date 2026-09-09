using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;
using Aspose.Words.Replacing;

public class HyphenationPdfExample
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow page width forces line wrapping and hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 200; // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Write a long word that can be hyphenated.
        builder.Font.Size = 24;
        builder.Writeln("extraordinarycharacteristically");

        // Create a minimal hyphenation dictionary for English (US).
        const string dictFileName = "hyph_en_US.dic";
        File.WriteAllText(dictFileName,
            "UTF-8\nextraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\n");

        // Register the dictionary.
        Hyphenation.RegisterDictionary("en-US", dictFileName);

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch

        // Recalculate layout so hyphenation is applied.
        doc.UpdatePageLayout();

        // Remove the hyphen characters inserted by hyphenation.
        doc.Range.Replace("-", string.Empty, new FindReplaceOptions());

        // Save the document as PDF.
        const string pdfFileName = "Hyphenated.pdf";
        doc.Save(pdfFileName, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(pdfFileName))
            throw new InvalidOperationException("Expected PDF output file was not created.");
    }
}
