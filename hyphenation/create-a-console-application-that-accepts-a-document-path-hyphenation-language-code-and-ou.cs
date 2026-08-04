using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main(string[] args)
    {
        // Input document path (first argument) or default name.
        string inputPath = args.Length > 0 && !string.IsNullOrWhiteSpace(args[0])
            ? args[0]
            : "sample.docx";

        // Hyphenation language code (second argument) or default "en-US".
        string language = args.Length > 1 && !string.IsNullOrWhiteSpace(args[1])
            ? args[1]
            : "en-US";

        // Output PDF path – same name as input but with .pdf extension.
        string outputPath = Path.ChangeExtension(inputPath, ".pdf");

        // Ensure the input document exists; if not, create a simple one with long words.
        if (!File.Exists(inputPath))
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("extraordinarycharacteristically internationalization communication");
            // Narrow the page width so hyphenation can occur.
            doc.FirstSection.PageSetup.PageWidth = 200;
            doc.FirstSection.PageSetup.LeftMargin = 20;
            doc.FirstSection.PageSetup.RightMargin = 20;
            doc.Save(inputPath);
        }

        // Create a minimal hyphenation dictionary for the requested language.
        // File name format: hyph_<language>.dic (replace characters that are invalid in file names).
        string safeLanguage = language.Replace('-', '_');
        string dictPath = $"hyph_{safeLanguage}.dic";

        if (!File.Exists(dictPath))
        {
            // Very small dictionary containing hyphenation patterns for the sample words.
            string dictContent = @"UTF-8
extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly
internationalization=in-ter-na-tion-al-i-za-tion
communication=com-mu-ni-ca-tion
";
            File.WriteAllText(dictPath, dictContent);
        }

        // Register the dictionary with Aspose.Words.
        Hyphenation.RegisterDictionary(language, dictPath);

        // Load the document.
        Document loadedDoc = new Document(inputPath);

        // Enable automatic hyphenation.
        loadedDoc.HyphenationOptions.AutoHyphenation = true;

        // Optionally set the locale of the document runs to match the language.
        // This helps Aspose.Words pick the correct hyphenation rules.
        int lcid = new CultureInfo(language).LCID;
        foreach (Run run in loadedDoc.GetChildNodes(NodeType.Run, true))
        {
            run.Font.LocaleId = lcid;
        }

        // Save as PDF.
        loadedDoc.Save(outputPath, SaveFormat.Pdf);

        // Validate that the PDF was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the output PDF at '{outputPath}'.");
    }
}
