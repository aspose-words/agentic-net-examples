using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class HyphenationExample
{
    public static void Main()
    {
        // Create a sample source DOCX.
        const string sourcePath = "source.docx";
        CreateSourceDocument(sourcePath);

        // Load the DOCX.
        Document doc = new Document(sourcePath);

        // Create and register a minimal hyphenation dictionary for English (US).
        const string dictPath = "hyph_en_US.dic";
        CreateDictionaryFile(dictPath);
        Hyphenation.RegisterDictionary("en-US", dictPath);

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.HyphenateCaps = true;

        // Save the document preserving hyphenation marks.
        const string outputPath = "hyphenated.docx";
        doc.Save(outputPath);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The hyphenated document was not created.");
    }

    private static void CreateSourceDocument(string path)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow the page to force line wrapping where hyphenation can occur.
        doc.FirstSection.PageSetup.PageWidth = 200;
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Add long words that will be hyphenated.
        builder.Font.Size = 24;
        builder.Writeln("extraordinarycharacteristically internationalization communication");

        doc.Save(path);
    }

    private static void CreateDictionaryFile(string path)
    {
        // Minimal OpenOffice hyphenation dictionary content.
        string content = @"UTF-8
extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly
internationalization=in-ter-na-tion-al-i-za-tion
communication=com-mu-ni-ca-tion
";
        File.WriteAllText(path, content);
    }
}
