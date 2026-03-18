using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;
using Aspose.Words.Saving;

class HyphenateBatchProcessor
{
    static void Main()
    {
        // Folder that contains the source DOCX files.
        string inputFolder = @"C:\Docs\Input";

        // Folder where the resulting PDF files will be written.
        string outputFolder = @"C:\Docs\Output";

        // Folder that holds hyphenation dictionary files (e.g., hyph_en-US.dic).
        string dictionariesFolder = @"C:\Docs\Dictionaries";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Iterate over every DOCX file in the input directory.
        foreach (string docxPath in Directory.GetFiles(inputFolder, "*.docx"))
        {
            // Load the DOCX document.
            Document doc = new Document(docxPath);

            // Determine the language code to use for hyphenation.
            // Expected file naming convention: <BaseName>_<lang>.docx (e.g., Report_en-US.docx).
            string language = ExtractLanguageFromFileName(Path.GetFileNameWithoutExtension(docxPath));

            // Register the appropriate hyphenation dictionary if it hasn't been registered yet.
            EnsureDictionaryRegistered(language, dictionariesFolder);

            // Enable automatic hyphenation and configure optional parameters.
            doc.HyphenationOptions.AutoHyphenation = true;
            doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;   // Max consecutive hyphenated lines.
            doc.HyphenationOptions.HyphenationZone = 720;       // 0.5 inch from right margin.
            doc.HyphenationOptions.HyphenateCaps = true;       // Hyphenate all‑caps words.

            // Construct the output PDF file path.
            string pdfFileName = Path.GetFileNameWithoutExtension(docxPath) + ".pdf";
            string pdfPath = Path.Combine(outputFolder, pdfFileName);

            // Save the document as PDF.
            doc.Save(pdfPath, SaveFormat.Pdf);
        }
    }

    // Extracts a language identifier from a file name.
    // If the name does not contain a language suffix, defaults to "en-US".
    private static string ExtractLanguageFromFileName(string fileNameWithoutExtension)
    {
        int underscoreIndex = fileNameWithoutExtension.LastIndexOf('_');
        if (underscoreIndex >= 0 && underscoreIndex < fileNameWithoutExtension.Length - 1)
        {
            string candidate = fileNameWithoutExtension.Substring(underscoreIndex + 1);
            if (candidate.Contains("-")) // Simple validation for culture format like "en-US".
                return candidate;
        }
        return "en-US";
    }

    // Registers a hyphenation dictionary for the specified language if needed.
    private static void EnsureDictionaryRegistered(string language, string dictionariesFolder)
    {
        // If a dictionary is already registered for this language, nothing to do.
        if (Hyphenation.IsDictionaryRegistered(language))
            return;

        // Expected dictionary file name pattern: hyph_<language>.dic
        string dictionaryFileName = $"hyph_{language}.dic";
        string dictionaryPath = Path.Combine(dictionariesFolder, dictionaryFileName);

        if (File.Exists(dictionaryPath))
        {
            // Register the dictionary from the file.
            Hyphenation.RegisterDictionary(language, dictionaryPath);
        }
        else
        {
            // No dictionary available – register a null dictionary to prevent repeated callbacks.
            Hyphenation.RegisterDictionary(language, (string)null);
        }
    }
}
