using System;
using System.IO;
using Aspose.Words;

class HyphenationDictionaryUpdater
{
    static void Main()
    {
        // Path to the original OpenOffice hyphenation dictionary.
        const string originalDictionaryPath = @"C:\Dictionaries\hyph_en_US.dic";

        // Path for the temporary dictionary that will contain the original patterns
        // plus the custom technical terminology patterns.
        string customDictionaryPath = Path.Combine(Path.GetTempPath(), "hyph_en_US_custom.dic");

        // Define the custom hyphenation patterns you want to add.
        // Each pattern must follow the OpenOffice dictionary format.
        string[] customPatterns =
        {
            "tech-no-lo-gy",   // example: "technology" can be hyphenated as tech-no-lo-gy
            "in-fra-struct-ure" // example: "infrastructure" as in-fra-struct-ure
        };

        // -----------------------------------------------------------------
        // Create a new dictionary file that merges the original patterns with the custom ones.
        // -----------------------------------------------------------------
        // Read all existing lines from the original dictionary.
        string[] originalLines = File.ReadAllLines(originalDictionaryPath);

        // Write the merged content to the temporary file.
        using (StreamWriter writer = new StreamWriter(customDictionaryPath, false))
        {
            // Preserve the original dictionary content.
            foreach (string line in originalLines)
                writer.WriteLine(line);

            // Append a comment to separate custom entries (optional).
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
            // The Hyphenation class lives directly under the Aspose.Words namespace in older versions.
            Aspose.Words.Hyphenation.RegisterDictionary("en-US", dictStream);
        }

        // -----------------------------------------------------------------
        // (Optional) Enable automatic hyphenation for a document that uses this language.
        // -----------------------------------------------------------------
        Document doc = new Document(); // Load or create a document as needed.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch from the right margin.
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenateCaps = true;

        // The document can now be saved using the standard Aspose.Words save logic.
        // Example (replace with your actual save path and format):
        // doc.Save(@"C:\Output\HyphenatedDocument.docx");
    }
}
