using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX document with words that can be hyphenated.
        var sourceDoc = new Document();
        var builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("extraordinarycharacteristically internationalization communication");

        // Narrow the page width to force line wrapping and hyphenation.
        sourceDoc.FirstSection.PageSetup.PageWidth = 200;
        sourceDoc.FirstSection.PageSetup.LeftMargin = 20;
        sourceDoc.FirstSection.PageSetup.RightMargin = 20;

        const string sourcePath = "sample.docx";
        sourceDoc.Save(sourcePath);
        if (!File.Exists(sourcePath))
            throw new InvalidOperationException("Source DOCX file was not created.");

        // Create a minimal Hunspell hyphenation dictionary.
        const string dictPath = "hyph_en_US.dic";
        var dictContent = @"UTF-8
extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly
internationalization=in-ter-na-tion-al-i-za-tion
communication=com-mu-ni-ca-tion
";
        File.WriteAllText(dictPath, dictContent);
        if (!File.Exists(dictPath))
            throw new InvalidOperationException("Hyphenation dictionary file was not created.");

        // Register the dictionary for the "en-US" locale.
        Hyphenation.RegisterDictionary("en-US", dictPath);

        // Load the previously saved document.
        var doc = new Document(sourcePath);

        // Enable automatic hyphenation for the loaded document.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Save the hyphenated document as PDF.
        const string outputPath = "hyphenated.pdf";
        doc.Save(outputPath);
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Hyphenated PDF file was not created.");
    }
}
