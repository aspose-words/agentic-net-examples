using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Create a minimal hyphenation dictionary for English (US).
        const string dictPath = "hyph_en_US.dic";
        File.WriteAllText(dictPath,
            "UTF-8\nextraordinarycharacteristically=ex-tra-or-di-nary-char-ac-ter-is-ti-cal-ly\ncommunication=com-mu-ni-ca-tion\n");

        // Register the dictionary for the "en-US" locale.
        Hyphenation.RegisterDictionary("en-US", dictPath);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow the page width to force line wrapping and possible hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 200;
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Add text containing a short word and longer words that can be hyphenated.
        // The short word "short" should not be hyphenated, while the longer words may be.
        builder.Writeln("short extraordinarycharacteristically communication");

        // Save the document to PDF.
        const string pdfPath = "HyphenationMinLength.pdf";
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Validate that the PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The PDF output file was not created.");

        // Load the PDF back as a document to ensure it can be read.
        Document loadedPdf = new Document(pdfPath);
        if (loadedPdf.PageCount == 0)
            throw new InvalidOperationException("The loaded PDF does not contain any pages.");

        // All validations passed.
        Console.WriteLine("Hyphenation example executed successfully.");
    }
}
