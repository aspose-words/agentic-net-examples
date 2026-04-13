using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class Program
{
    // Callback that registers hyphenation dictionaries on demand.
    private class HyphenationCallback : IHyphenationCallback
    {
        private readonly Dictionary<string, string> _dictionaryFiles;

        public HyphenationCallback(Dictionary<string, string> dictionaryFiles)
        {
            _dictionaryFiles = dictionaryFiles;
        }

        public void RequestDictionary(string language)
        {
            Console.WriteLine($"Hyphenation dictionary requested for language: {language}");

            if (Hyphenation.IsDictionaryRegistered(language))
            {
                Console.WriteLine("  Dictionary already registered.");
                return;
            }

            if (_dictionaryFiles.TryGetValue(language, out string filePath) && File.Exists(filePath))
            {
                Hyphenation.RegisterDictionary(language, filePath);
                Console.WriteLine("  Dictionary registered successfully.");
            }
            else
            {
                Console.WriteLine("  No dictionary file found for the requested language.");
            }
        }
    }

    public static void Main()
    {
        // -----------------------------------------------------------------
        // Simulate a CI step that ensures the hyphenation dictionary is up‑to‑date.
        // -----------------------------------------------------------------
        string baseDir = Directory.GetCurrentDirectory();
        string dictDir = Path.Combine(baseDir, "HyphenationDictionaries");
        Directory.CreateDirectory(dictDir);

        // Create a minimal placeholder dictionary for English (US).
        string enUsDictPath = Path.Combine(dictDir, "hyph_en_US.dic");
        if (!File.Exists(enUsDictPath))
        {
            // A single empty pattern is sufficient for the example.
            File.WriteAllText(enUsDictPath, "1\n");
            Console.WriteLine($"Created placeholder dictionary: {enUsDictPath}");
        }

        // Register the dictionary directly (optional – the callback will also register it on demand).
        if (!Hyphenation.IsDictionaryRegistered("en-US"))
        {
            Hyphenation.RegisterDictionary("en-US", enUsDictPath);
            Console.WriteLine("Dictionary registered via direct call.");
        }

        // -----------------------------------------------------------------
        // Prepare the callback that can register dictionaries on demand.
        // -----------------------------------------------------------------
        var dictionaryMap = new Dictionary<string, string>
        {
            { "en-US", enUsDictPath }
        };
        Hyphenation.Callback = new HyphenationCallback(dictionaryMap);

        // -----------------------------------------------------------------
        // Create a sample document that will demonstrate hyphenation.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Use a narrow page width to force line wrapping.
        doc.FirstSection.PageSetup.PageWidth = 200; // points (~2.78 inches)
        doc.FirstSection.PageSetup.PageHeight = 842; // A4 height in points (optional)

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;
        // Set a positive hyphenation zone (default is 360 = 0.25 inch).
        doc.HyphenationOptions.HyphenationZone = 360;

        // Set the locale for the paragraph to English (US) so the correct dictionary is used.
        builder.Font.LocaleId = new CultureInfo("en-US").LCID;
        builder.Font.Size = 24;

        // A long sentence that can be hyphenated.
        builder.Writeln("Antidisestablishmentarianism is a long word that may be hyphenated across lines.");

        // -----------------------------------------------------------------
        // Save the document to PDF – this triggers layout and hyphenation.
        // -----------------------------------------------------------------
        string outputPath = Path.Combine(baseDir, "HyphenatedDocument.pdf");
        doc.Save(outputPath, SaveFormat.Pdf);
        Console.WriteLine($"Document saved to: {outputPath}");

        // -----------------------------------------------------------------
        // Validation – ensure the output file exists.
        // -----------------------------------------------------------------
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("Failed to create the PDF output file.");
        }

        Console.WriteLine("Hyphenation processing completed successfully.");
    }
}
