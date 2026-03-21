using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

class Program
{
    static void Main()
    {
        // Prepare folders relative to the executable directory.
        string baseDir = AppContext.BaseDirectory;
        string dataDir = Path.Combine(baseDir, "Data");
        string artifactsDir = Path.Combine(baseDir, "Artifacts");

        Directory.CreateDirectory(dataDir);
        Directory.CreateDirectory(artifactsDir);

        // Ensure a (dummy) hyphenation dictionary exists so RegisterDictionary does not throw.
        string germanDictPath = Path.Combine(dataDir, "hyph_de_CH.dic");
        if (!File.Exists(germanDictPath))
        {
            // A minimal dictionary file – real hyphenation patterns are not required for this demo.
            File.WriteAllText(germanDictPath, "%% Dummy hyphenation dictionary");
        }

        // Register the German hyphenation dictionary (de-CH) from the file.
        try
        {
            Hyphenation.RegisterDictionary("de-CH", germanDictPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Warning: could not register hyphenation dictionary: {ex.Message}");
        }

        // Optional: set a callback to load dictionaries on demand.
        Hyphenation.Callback = new HyphenationCallback(dataDir);

        // Create a simple German language document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Dies ist ein Beispieltext, der die Silbentrennung in deutscher Sprache demonstriert.");

        // Enable automatic hyphenation for the document.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.HyphenateCaps = true;
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch (720 / 1440 points)

        // Save the hyphenated document.
        string outputPath = Path.Combine(artifactsDir, "GermanHyphenated.docx");
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");
    }

    // Callback that registers hyphenation dictionaries when requested by the layout engine.
    private class HyphenationCallback : IHyphenationCallback
    {
        private readonly Dictionary<string, string> _dictionaryFiles;

        public HyphenationCallback(string basePath)
        {
            _dictionaryFiles = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "de-CH", Path.Combine(basePath, "hyph_de_CH.dic") },
                { "en-US", Path.Combine(basePath, "hyph_en_US.dic") }
            };
        }

        public void RequestDictionary(string language)
        {
            if (Hyphenation.IsDictionaryRegistered(language))
                return;

            if (_dictionaryFiles.TryGetValue(language, out string filePath) && File.Exists(filePath))
                Hyphenation.RegisterDictionary(language, filePath);
        }
    }
}
