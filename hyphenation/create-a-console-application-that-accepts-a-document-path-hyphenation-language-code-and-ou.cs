using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class HyphenationPdfGenerator
{
    public static void Main(string[] args)
    {
        // Default values when arguments are not supplied.
        string inputPath = args.Length > 0 ? args[0] : "sample.docx";
        string languageCode = args.Length > 1 ? args[1] : "en-US";
        string outputPath = args.Length > 2 ? args[2] : Path.ChangeExtension(inputPath, ".pdf");

        // Ensure the input document exists; if not, create a simple sample document.
        if (!File.Exists(inputPath))
        {
            CreateSampleDocument(inputPath, languageCode);
        }

        // Create a minimal hyphenation dictionary for the requested language.
        string dictionaryPath = $"hyph_{languageCode.Replace("-", "_")}.dic";
        if (!File.Exists(dictionaryPath))
        {
            CreateMinimalDictionary(dictionaryPath);
        }

        // Register the dictionary so Aspose.Words can hyphenate words of this language.
        Hyphenation.RegisterDictionary(languageCode, dictionaryPath);

        // Load the document.
        Document doc = new Document(inputPath);

        // If the document has no paragraphs, add a sample paragraph to demonstrate hyphenation.
        if (doc.FirstSection?.Body?.Paragraphs?.Count == 0)
        {
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("extraordinarycharacteristically internationalization communication");
        }

        // Apply the language locale to all runs in the document.
        foreach (Run run in doc.GetChildNodes(NodeType.Run, true))
        {
            run.Font.LocaleId = new CultureInfo(languageCode).LCID;
        }

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Narrow the page width to force line wrapping where hyphenation can occur.
        doc.FirstSection.PageSetup.PageWidth = 200;
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Save the document as PDF.
        doc.Save(outputPath, SaveFormat.Pdf);

        // Validate that the PDF was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The expected PDF output file was not created.");
    }

    // Creates a simple DOCX document containing text that can be hyphenated.
    private static void CreateSampleDocument(string path, string languageCode)
    {
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Font.Size = 12;
        builder.Writeln("extraordinarycharacteristically internationalization communication");

        // Set the locale for the sample text.
        foreach (Run run in sampleDoc.GetChildNodes(NodeType.Run, true))
        {
            run.Font.LocaleId = new CultureInfo(languageCode).LCID;
        }

        sampleDoc.Save(path);
    }

    // Writes a minimal hyphenation dictionary file in OpenOffice format.
    private static void CreateMinimalDictionary(string dictPath)
    {
        // The dictionary must start with the encoding line.
        // Follow with a few sample words and their hyphenation patterns.
        string content = @"UTF-8
extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly
internationalization=in-ter-na-tion-al-i-za-tion
communication=com-mu-ni-ca-tion
";
        File.WriteAllText(dictPath, content);
    }
}
