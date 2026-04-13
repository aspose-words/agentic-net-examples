using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define a deterministic folder for temporary files.
        string dataDir = Path.Combine(Directory.GetCurrentDirectory(), "Data");
        Directory.CreateDirectory(dataDir);

        // Create a simple French hyphenation dictionary file.
        // The OpenOffice hyphenation format consists of pattern lines.
        string frenchDictPath = Path.Combine(dataDir, "hyph_fr_FR.dic");
        string[] samplePatterns =
        {
            "%% French hyphenation patterns (sample)",
            "1ba 1be 1bi 1bo 1bu",
            "2c3h 2g3h 2ph 2th",
            "3é 3è 3ê 3ë"
        };
        File.WriteAllLines(frenchDictPath, samplePatterns);

        // Register the French dictionary with Aspose.Words.
        // Language code follows .NET culture naming (e.g., "fr-FR").
        Hyphenation.RegisterDictionary("fr-FR", frenchDictPath);

        // Verify that the dictionary is registered.
        bool isRegistered = Hyphenation.IsDictionaryRegistered("fr-FR");
        Console.WriteLine($"French hyphenation dictionary registered: {isRegistered}");

        // For debugging, read back the dictionary file content (the patterns).
        Console.WriteLine("Hyphenation patterns for French (fr-FR):");
        foreach (string line in File.ReadAllLines(frenchDictPath))
        {
            Console.WriteLine(line);
        }

        // Clean up: optionally delete the temporary files.
        // Comment out the following lines if you want to inspect the files after execution.
        try
        {
            File.Delete(frenchDictPath);
            Directory.Delete(dataDir);
        }
        catch
        {
            // Ignored – cleanup is best‑effort.
        }
    }
}
