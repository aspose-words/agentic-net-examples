using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class HyphenateBatchProcessor
{
    static void Main()
    {
        // Use folders relative to the executable location.
        string baseDir = AppContext.BaseDirectory;
        string inputFolder = Path.Combine(baseDir, "Input");
        string outputFolder = Path.Combine(baseDir, "Output");
        string dictionariesFolder = Path.Combine(baseDir, "Dictionaries");

        // Ensure required directories exist.
        Directory.CreateDirectory(inputFolder);
        Directory.CreateDirectory(outputFolder);
        Directory.CreateDirectory(dictionariesFolder);

        // If there are no DOCX files, inform the user and exit gracefully.
        var docxFiles = Directory.EnumerateFiles(inputFolder, "*.docx");
        bool anyFile = false;
        foreach (var docxPath in docxFiles)
        {
            anyFile = true;
            ProcessDocument(docxPath, outputFolder, dictionariesFolder);
        }

        if (!anyFile)
        {
            Console.WriteLine($"No DOCX files found in '{inputFolder}'. Place files there and rerun the program.");
        }
    }

    private static void ProcessDocument(string docxPath, string outputFolder, string dictionariesFolder)
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

        Console.WriteLine($"Processed '{Path.GetFileName(docxPath)}' → '{pdfFileName}'.");
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
