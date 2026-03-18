using System;
using System.IO;
using System.Diagnostics;
using Aspose.Words;
using Aspose.Words.Loading;

class HyphenationDebug
{
    static void Main()
    {
        // Path to the folder that contains the hyphenation dictionary files.
        // Adjust this path to point to the location of your French hyphenation dictionary (hyph_fr_FR.dic).
        string myDir = @"C:\HyphenationDictionaries";

        // Full path to the French hyphenation dictionary file.
        string frenchDicPath = Path.Combine(myDir, "hyph_fr_FR.dic");

        // Verify that the dictionary file exists.
        if (!File.Exists(frenchDicPath))
        {
            Console.WriteLine($"French hyphenation dictionary not found at: {frenchDicPath}");
            return;
        }

        // Register the French dictionary so that Aspose.Words can use it during layout.
        // This also loads the patterns into the internal hyphenation engine.
        using (FileStream stream = new FileStream(frenchDicPath, FileMode.Open, FileAccess.Read))
        {
            Hyphenation.RegisterDictionary("fr-FR", stream);
        }

        // Confirm registration.
        bool isRegistered = Hyphenation.IsDictionaryRegistered("fr-FR");
        Console.WriteLine($"French hyphenation dictionary registered: {isRegistered}");

        // Read and log the raw hyphenation patterns for debugging.
        // The dictionary file is a plain‑text OpenOffice format; each line is a pattern or a comment.
        Console.WriteLine("=== French Hyphenation Patterns ===");
        foreach (string line in File.ReadLines(frenchDicPath))
        {
            // Skip empty lines and comments (lines starting with '#').
            if (string.IsNullOrWhiteSpace(line) || line.StartsWith("#"))
                continue;

            Debug.WriteLine(line);          // Sends the pattern to the debugger output.
            Console.WriteLine(line);        // Also writes to the console for visibility.
        }
        Console.WriteLine("=== End of Patterns ===");
    }
}
