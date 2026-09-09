using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class HyphenationExample
{
    public static void Main()
    {
        // Create a sample document with long words that can be hyphenated.
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("extraordinarycharacteristically internationalization communication");
        // Narrow the page to force line wrapping.
        sampleDoc.FirstSection.PageSetup.PageWidth = 200;
        sampleDoc.FirstSection.PageSetup.LeftMargin = 20;
        sampleDoc.FirstSection.PageSetup.RightMargin = 20;
        const string inputPath = "input.docx";
        sampleDoc.Save(inputPath);

        // Create a minimal Hunspell dictionary file for English (US).
        const string dictPath = "hyph_en_US.dic";
        string dictContent = @"UTF-8
extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly
internationalization=in-ter-na-tion-al-i-za-tion
communication=com-mu-ni-ca-tion
";
        File.WriteAllText(dictPath, dictContent);

        // Register the dictionary with Aspose.Words.
        Hyphenation.RegisterDictionary("en-US", dictPath);

        // Load the previously saved document.
        Document doc = new Document(inputPath);

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
            throw new InvalidOperationException("The hyphenated PDF was not created.");
    }
}
