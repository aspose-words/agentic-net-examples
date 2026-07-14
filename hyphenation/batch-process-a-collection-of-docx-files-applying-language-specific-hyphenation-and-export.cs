using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Settings;

public class Program
{
    // Entry point of the console application.
    public static void Main()
    {
        // Prepare working directories.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "OutputPdfs");
        string dictDir = Path.Combine(baseDir, "HyphenationDictionaries");

        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);
        Directory.CreateDirectory(dictDir);

        // Create minimal hyphenation dictionaries for English (US) and German (Switzerland).
        CreateDictionary(dictDir, "en-US", @"UTF-8
extraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly
internationalization=in-ter-na-tion-al-i-za-tion
communication=com-mu-ni-ca-tion
");

        CreateDictionary(dictDir, "de-CH", @"UTF-8
aussergewoehnlich=aus-ser-ge-woehn-lich
internationalisierung=in-ter-na-tion-a-li-sie-rung
kommunikation=ko-mmu-ni-ka-tion
");

        // Create sample DOCX files with language‑specific content.
        CreateSampleDocument(Path.Combine(inputDir, "sample_en.docx"), "en-US",
            "extraordinarycharacteristically internationalization communication extraordinarycharacteristically internationalization communication");

        CreateSampleDocument(Path.Combine(inputDir, "sample_de.docx"), "de-CH",
            "aussergewoehnlich internationalisierung kommunikation aussergewoehnlich internationalisierung kommunikation");

        // Process each DOCX file: apply hyphenation and export to PDF.
        foreach (string docxPath in Directory.GetFiles(inputDir, "*.docx"))
        {
            // Load the document.
            Document doc = new Document(docxPath);

            // Determine the language of the document from the first run's locale.
            string language = DetectDocumentLanguage(doc) ?? "en-US";

            // Register the appropriate hyphenation dictionary if not already registered.
            string dictPath = Path.Combine(dictDir, $"hyph_{language.Replace("-", "_")}.dic");
            if (!Hyphenation.IsDictionaryRegistered(language) && File.Exists(dictPath))
                Hyphenation.RegisterDictionary(language, dictPath);

            // Enable automatic hyphenation.
            doc.HyphenationOptions.AutoHyphenation = true;
            doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
            doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch
            doc.HyphenationOptions.HyphenateCaps = true;

            // Save as PDF.
            string pdfFileName = Path.GetFileNameWithoutExtension(docxPath) + ".pdf";
            string pdfPath = Path.Combine(outputDir, pdfFileName);
            doc.Save(pdfPath, SaveFormat.Pdf);

            // Validate that the PDF was created.
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"Failed to create PDF: {pdfPath}");
        }

        // Clean up (optional): uncomment the following lines if you want to remove temporary files after execution.
        //Directory.Delete(inputDir, true);
        //Directory.Delete(dictDir, true);
        //Directory.Delete(outputDir, true);
    }

    // Creates a hyphenation dictionary file with the given language code and content.
    private static void CreateDictionary(string directory, string language, string content)
    {
        string fileName = $"hyph_{language.Replace("-", "_")}.dic";
        string fullPath = Path.Combine(directory, fileName);
        File.WriteAllText(fullPath, content);
    }

    // Generates a sample DOCX file with the specified language locale and sample text.
    private static void CreateSampleDocument(string filePath, string language, string sampleText)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow page width to force line wrapping and hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 300; // points (~4.17 inches)
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Set the locale for the paragraph.
        builder.Font.LocaleId = new CultureInfo(language).LCID;
        builder.Font.Size = 12;
        builder.Writeln(sampleText);

        doc.Save(filePath);
    }

    // Attempts to detect the document language from the first run's locale.
    private static string DetectDocumentLanguage(Document doc)
    {
        foreach (Run run in doc.GetChildNodes(NodeType.Run, true))
        {
            if (run.Font.LocaleId != 0)
                return CultureInfo.GetCultureInfo(run.Font.LocaleId).Name;
        }
        return null;
    }
}
