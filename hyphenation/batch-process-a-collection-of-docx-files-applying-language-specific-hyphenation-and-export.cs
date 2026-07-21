using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare folders.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "OutputPdfs");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // Create minimal hyphenation dictionaries.
        CreateDictionary("en-US", "hyph_en_US.dic",
            "UTF-8\nextraordinarycharacteristically=extra-or-di-nary-char-ac-ter-is-ti-cal-ly\ninternationalization=in-ter-na-tion-al-i-za-tion\ncommunication=com-mu-ni-ca-tion\n");
        CreateDictionary("de-CH", "hyph_de_CH.dic",
            "UTF-8\ninternationalisierung=inter-na-tion-ali-sier-ung\nkommunikation=ko-mmu-ni-ka-tion\n");

        // Register dictionaries once (they will be reused for all documents).
        Aspose.Words.Hyphenation.RegisterDictionary("en-US", Path.Combine(baseDir, "hyph_en_US.dic"));
        Aspose.Words.Hyphenation.RegisterDictionary("de-CH", Path.Combine(baseDir, "hyph_de_CH.dic"));

        // Create sample DOCX files for the batch.
        CreateSampleDocument(Path.Combine(inputDir, "Sample_en-US.docx"),
            "extraordinarycharacteristically internationalization communication", "en-US");
        CreateSampleDocument(Path.Combine(inputDir, "Sample_de-CH.docx"),
            "internationalisierung kommunikation", "de-CH");

        // Process each DOCX file: apply hyphenation and export to PDF.
        foreach (string docxPath in Directory.GetFiles(inputDir, "*.docx"))
        {
            // Load the document.
            Document doc = new Document(docxPath);

            // Determine language code from file name (e.g., Sample_en-US.docx).
            string language = GetLanguageFromFileName(docxPath);
            if (string.IsNullOrEmpty(language))
                throw new InvalidOperationException($"Cannot determine language for file '{docxPath}'.");

            // Ensure the dictionary for this language is registered.
            if (!Aspose.Words.Hyphenation.IsDictionaryRegistered(language))
                throw new InvalidOperationException($"Hyphenation dictionary for language '{language}' is not registered.");

            // Enable automatic hyphenation.
            doc.HyphenationOptions.AutoHyphenation = true;

            // Save as PDF.
            string pdfPath = Path.Combine(outputDir, Path.GetFileNameWithoutExtension(docxPath) + ".pdf");
            doc.Save(pdfPath, SaveFormat.Pdf);

            // Validate output.
            if (!File.Exists(pdfPath))
                throw new InvalidOperationException($"Failed to create PDF '{pdfPath}'.");
        }

        // Indicate success.
        Console.WriteLine("Batch hyphenation and PDF conversion completed successfully.");
    }

    // Creates a simple hyphenation dictionary file.
    private static void CreateDictionary(string language, string fileName, string content)
    {
        string fullPath = Path.Combine(Directory.GetCurrentDirectory(), fileName);
        File.WriteAllText(fullPath, content);
        if (!File.Exists(fullPath))
            throw new InvalidOperationException($"Failed to create dictionary file '{fullPath}'.");
    }

    // Generates a sample DOCX with given text and locale.
    private static void CreateSampleDocument(string filePath, string text, string locale)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set a narrow page width to force line wrapping and hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 300;
        doc.FirstSection.PageSetup.LeftMargin = 20;
        doc.FirstSection.PageSetup.RightMargin = 20;

        // Apply locale to the paragraph.
        builder.Font.LocaleId = new CultureInfo(locale).LCID;
        builder.Writeln(text);

        doc.Save(filePath);
        if (!File.Exists(filePath))
            throw new InvalidOperationException($"Failed to create sample document '{filePath}'.");
    }

    // Extracts the language code from a file name like "Sample_en-US.docx".
    private static string GetLanguageFromFileName(string filePath)
    {
        string fileName = Path.GetFileNameWithoutExtension(filePath);
        int underscoreIndex = fileName.LastIndexOf('_');
        if (underscoreIndex < 0 || underscoreIndex == fileName.Length - 1)
            return string.Empty;
        return fileName.Substring(underscoreIndex + 1);
    }
}
