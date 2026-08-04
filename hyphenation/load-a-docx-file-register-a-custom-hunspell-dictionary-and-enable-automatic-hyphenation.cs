using System;
using System.Globalization;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define file names.
        const string inputDocx = "sample.docx";
        const string hyphenDict = "hyph_en_US.dic";
        const string outputPdf = "hyphenated.pdf";

        // -----------------------------------------------------------------
        // 1. Create a sample DOCX file with long words that can be hyphenated.
        // -----------------------------------------------------------------
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Narrow the page width to force line wrapping.
        doc.FirstSection.PageSetup.PageWidth = 200;   // points
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Write sample text containing words we will define in the dictionary.
        builder.Font.Size = 12;
        builder.Writeln("extraordinarycharacteristically internationalization communication");

        // Save the sample document.
        doc.Save(inputDocx);

        // -----------------------------------------------------------------
        // 2. Create a minimal Hunspell (OpenOffice) hyphenation dictionary.
        // -----------------------------------------------------------------
        // The dictionary format: first line is the encoding, followed by word=hyphenation-pattern.
        var dictContent = @"UTF-8
extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly
internationalization=in-ter-na-tion-al-i-za-tion
communication=com-mu-ni-ca-tion
";
        File.WriteAllText(hyphenDict, dictContent);

        // -----------------------------------------------------------------
        // 3. Load the DOCX file, register the dictionary, and enable hyphenation.
        // -----------------------------------------------------------------
        var loadedDoc = new Document(inputDocx);

        // Register the dictionary for the "en-US" locale.
        Aspose.Words.Hyphenation.RegisterDictionary("en-US", hyphenDict);

        // Verify registration.
        if (!Aspose.Words.Hyphenation.IsDictionaryRegistered("en-US"))
            throw new InvalidOperationException("Hyphenation dictionary registration failed.");

        // Enable automatic hyphenation for the document.
        loadedDoc.HyphenationOptions.AutoHyphenation = true;

        // Ensure the paragraph uses the matching locale.
        var firstParagraph = loadedDoc.FirstSection.Body.FirstParagraph;
        firstParagraph.ParagraphFormat.SuppressAutoHyphens = false;
        if (firstParagraph.Runs.Count > 0)
            firstParagraph.Runs[0].Font.LocaleId = new CultureInfo("en-US").LCID;

        // -----------------------------------------------------------------
        // 4. Save the hyphenated document to PDF.
        // -----------------------------------------------------------------
        loadedDoc.Save(outputPdf);

        // Validate that the output file was created.
        if (!File.Exists(outputPdf))
            throw new InvalidOperationException($"Expected output file '{outputPdf}' was not created.");
    }
}
