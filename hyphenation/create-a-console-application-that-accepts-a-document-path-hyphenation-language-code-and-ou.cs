using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class HyphenationExample
{
    public static void Main(string[] args)
    {
        // Resolve arguments or use defaults.
        string documentPath = args.Length > 0 ? args[0] : "sample.docx";
        string languageCode = args.Length > 1 ? args[1] : "en-US";

        // Ensure a source document exists.
        if (!File.Exists(documentPath))
        {
            CreateSampleDocument(documentPath, languageCode);
        }

        // Create a minimal hyphenation dictionary for the requested language.
        string dictionaryPath = $"hyph_{languageCode}.dic";
        CreateDictionaryFile(dictionaryPath, languageCode);

        // Register the dictionary with Aspose.Words.
        Hyphenation.RegisterDictionary(languageCode, dictionaryPath);

        // Load the document and enable hyphenation.
        Document doc = new Document(documentPath);
        doc.HyphenationOptions.AutoHyphenation = true;

        // Narrow the page width to force line breaks where hyphenation can occur.
        Section section = doc.FirstSection;
        section.PageSetup.PageWidth = 300; // points
        section.PageSetup.LeftMargin = 20;
        section.PageSetup.RightMargin = 20;

        // Save as PDF.
        string pdfPath = Path.ChangeExtension(documentPath, ".pdf");
        doc.Save(pdfPath, SaveFormat.Pdf);

        // Validate output.
        if (!File.Exists(pdfPath))
        {
            throw new InvalidOperationException($"PDF output was not created at '{pdfPath}'.");
        }
    }

    private static void CreateSampleDocument(string path, string languageCode)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set the document locale to match the hyphenation language.
        builder.Font.LocaleId = new CultureInfo(languageCode).LCID;

        // Write a paragraph containing words that will be hyphenated.
        builder.Writeln("extraordinarycharacteristically internationalization communication");
        builder.Writeln("hyperresponsibility uncharacteristically misunderstanding");

        // Save the document.
        doc.Save(path, SaveFormat.Docx);
    }

    private static void CreateDictionaryFile(string path, string languageCode)
    {
        // Minimal dictionary content: first line is the encoding, followed by word=hyphenated-form.
        string[] lines =
        {
            "UTF-8",
            "extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly",
            "internationalization=in-ter-na-tion-al-i-za-tion",
            "communication=com-mu-ni-ca-tion",
            "hyperresponsibility=hy-per-re-spon-si-bil-i-ty",
            "uncharacteristically=un-char-ac-ter-is-ti-cal-ly",
            "misunderstanding=mis-un-der-stand-ing"
        };
        File.WriteAllLines(path, lines);
    }
}
