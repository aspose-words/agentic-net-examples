using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Base directory of the application.
        string baseDir = AppDomain.CurrentDomain.BaseDirectory;

        // Directories for input documents, output PDFs and hyphenation dictionaries.
        string inputDir = Path.Combine(baseDir, "InputDocs");
        string outputDir = Path.Combine(baseDir, "OutputPdfs");
        string dictDir = Path.Combine(baseDir, "HyphenationDictionaries");

        // Ensure directories exist.
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);
        Directory.CreateDirectory(dictDir);

        // Create minimal hyphenation dictionary files.
        CreateDictionaryFile(dictDir, "hyph_en_US.dic", "1\na1");
        CreateDictionaryFile(dictDir, "hyph_de_CH.dic", "1\ne1");

        // Create sample DOCX files with different locales.
        CreateSampleDocument(Path.Combine(inputDir, "English.docx"), "en-US",
            "This is a sample English text that should demonstrate hyphenation when the line width is narrow. " +
            "The quick brown fox jumps over the lazy dog.");

        CreateSampleDocument(Path.Combine(inputDir, "German.docx"), "de-CH",
            "Dies ist ein Beispieltext auf Deutsch, um die Silbentrennung zu demonstrieren, wenn die Zeilenbreite schmal ist. " +
            "Der schnelle braune Fuchs springt über den faulen Hund.");

        // Register the callback that will load dictionaries on demand.
        Hyphenation.Callback = new CustomHyphenationDictionaryRegister(dictDir);

        // Process each DOCX file: enable hyphenation and export to PDF.
        foreach (string docxPath in Directory.GetFiles(inputDir, "*.docx"))
        {
            Document doc = new Document(docxPath);

            // Enable automatic hyphenation.
            doc.HyphenationOptions.AutoHyphenation = true;
            doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
            doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch.

            // Narrow the page width to force line wrapping and hyphenation.
            doc.FirstSection.PageSetup.PageWidth = 300; // Points (~4.17 inches).

            string pdfPath = Path.Combine(outputDir, Path.GetFileNameWithoutExtension(docxPath) + ".pdf");
            doc.Save(pdfPath, SaveFormat.Pdf);
        }

        // Validate that PDFs were created.
        foreach (string pdfPath in Directory.GetFiles(outputDir, "*.pdf"))
        {
            if (!File.Exists(pdfPath))
                throw new FileNotFoundException("Expected PDF file was not created.", pdfPath);
        }

        Console.WriteLine("Batch hyphenation and PDF conversion completed successfully.");
    }

    // Helper to create a simple hyphenation dictionary file.
    private static void CreateDictionaryFile(string directory, string fileName, string content)
    {
        string path = Path.Combine(directory, fileName);
        File.WriteAllText(path, content);
    }

    // Helper to create a sample DOCX with specified locale and text.
    private static void CreateSampleDocument(string filePath, string cultureName, string text)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Set a readable font size.
        builder.Font.Size = 24;

        // Apply the locale to the paragraph.
        builder.Font.LocaleId = new CultureInfo(cultureName).LCID;

        // Write the sample text.
        builder.Writeln(text);

        // Save the document.
        doc.Save(filePath);
    }

    // Callback implementation that registers dictionaries from the local folder.
    private class CustomHyphenationDictionaryRegister : IHyphenationCallback
    {
        private readonly Dictionary<string, string> _dictionaryFiles;

        public CustomHyphenationDictionaryRegister(string dictionaryDirectory)
        {
            _dictionaryFiles = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "en-US", Path.Combine(dictionaryDirectory, "hyph_en_US.dic") },
                { "de-CH", Path.Combine(dictionaryDirectory, "hyph_de_CH.dic") }
            };
        }

        public void RequestDictionary(string language)
        {
            // If already registered, nothing to do.
            if (Hyphenation.IsDictionaryRegistered(language))
                return;

            // Register the dictionary if we have a matching file.
            if (_dictionaryFiles.TryGetValue(language, out string filePath) && File.Exists(filePath))
            {
                Hyphenation.RegisterDictionary(language, filePath);
                return;
            }

            // If no dictionary is available, register a null dictionary to suppress further callbacks.
            Hyphenation.RegisterDictionary(language, (Stream)null);
        }
    }
}
