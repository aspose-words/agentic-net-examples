using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using static Aspose.Words.Hyphenation; // Import static members of Hyphenation

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX with long words that can be hyphenated.
        var sourceDoc = new Document();
        var builder = new DocumentBuilder(sourceDoc);
        builder.Font.Size = 24;
        builder.Writeln("extraordinarycharacteristically internationalization communication");

        // Narrow page width forces line wrapping.
        sourceDoc.FirstSection.PageSetup.PageWidth = 200;
        sourceDoc.FirstSection.PageSetup.LeftMargin = 20;
        sourceDoc.FirstSection.PageSetup.RightMargin = 20;

        const string sourcePath = "source.docx";
        sourceDoc.Save(sourcePath); // Save the source document.

        // Create a minimal Hunspell dictionary for English (US).
        const string dictPath = "hyph_en_US.dic";
        const string dictContent = @"UTF-8
extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly
internationalization=in-ter-na-tion-al-i-za-tion
communication=com-mu-ni-ca-tion
";
        File.WriteAllText(dictPath, dictContent);

        // Register the dictionary for the "en-US" locale.
        RegisterDictionary("en-US", dictPath);

        // Load the previously saved DOCX.
        var doc = new Document(sourcePath);

        // Ensure the document uses the same locale as the registered dictionary.
        if (doc.FirstSection?.Body?.FirstParagraph?.Runs?.Count > 0)
        {
            doc.FirstSection.Body.FirstParagraph.Runs[0].Font.LocaleId = new CultureInfo("en-US").LCID;
        }

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.HyphenateCaps = true;
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenationZone = 720;

        // Save the hyphenated document to PDF.
        const string outputPath = "hyphenated.pdf";
        doc.Save(outputPath);

        // Validate that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Expected output file was not created.");
    }
}
