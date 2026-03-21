using System;
using System.IO;
using Aspose.Words;

class HyphenationDictionaryUpdater
{
    static void Main()
    {
        // Use a temporary location for the original dictionary so the example works out‑of‑the‑box.
        string originalDictionaryPath = Path.Combine(Path.GetTempPath(), "hyph_en_US.dic");

        // If the original dictionary does not exist, create a minimal placeholder.
        if (!File.Exists(originalDictionaryPath))
        {
            File.WriteAllLines(originalDictionaryPath, new[]
            {
                "# Minimal placeholder OpenOffice hyphenation dictionary",
                "a-bc",   // example pattern
                "de-fg"
            });
        }

        // Path for the temporary dictionary that will contain the original patterns
        // plus the custom technical terminology patterns.
        string customDictionaryPath = Path.Combine(Path.GetTempPath(), "hyph_en_US_custom.dic");

        // Define the custom hyphenation patterns you want to add.
        // Each pattern must follow the OpenOffice dictionary format.
        string[] customPatterns =
        {
            "tech-no-lo-gy",       // "technology" → tech-no-lo-gy
            "in-fra-struct-ure"    // "infrastructure" → in-fra-struct-ure
        };

        // -----------------------------------------------------------------
        // Create a new dictionary file that merges the original patterns with the custom ones.
        // -----------------------------------------------------------------
        string[] originalLines = File.ReadAllLines(originalDictionaryPath);

        using (var writer = new StreamWriter(customDictionaryPath, false))
        {
            // Preserve the original dictionary content.
            foreach (string line in originalLines)
                writer.WriteLine(line);

            // Optional separator comment.
            writer.WriteLine("# Custom technical terminology patterns");

            // Append each custom pattern.
            foreach (string pattern in customPatterns)
                writer.WriteLine(pattern);
        }

        // -----------------------------------------------------------------
        // Register the newly created dictionary with Aspose.Words.
        // -----------------------------------------------------------------
        using (FileStream dictStream = new FileStream(customDictionaryPath, FileMode.Open, FileAccess.Read))
        {
            // Register the dictionary for the desired language (e.g., en-US).
            Hyphenation.RegisterDictionary("en-US", dictStream);
        }

        // -----------------------------------------------------------------
        // (Optional) Enable automatic hyphenation for a document that uses this language.
        // -----------------------------------------------------------------
        Document doc = new Document(); // Load or create a document as needed.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch from the right margin.
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenateCaps = true;

        // Example save (adjust path as needed):
        // doc.Save(@"C:\Output\HyphenatedDocument.docx");
    }
}
