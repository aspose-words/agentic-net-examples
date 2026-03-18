using System;
using System.IO;
using Aspose.Words;
// The WarningInfoCollection class resides in the Aspose.Words namespace in older
// versions of the library (pre‑21.9). If you are using a newer version that still
// provides Aspose.Words.Warnings, you can keep the original using directive.
// For maximum compatibility we reference the root namespace only.

public class HyphenationErrorHandling
{
    /// <summary>
    /// Attempts to register a hyphenation dictionary and handles possible errors.
    /// </summary>
    /// <param name="language">Culture name, e.g. "en-US".</param>
    /// <param name="dictionaryPath">Full path to the OpenOffice‑format dictionary file.</param>
    public static void RegisterDictionaryWithHandling(string language, string dictionaryPath)
    {
        // Collect warnings that may be raised while loading the dictionary.
        WarningInfoCollection warnings = new WarningInfoCollection();
        Hyphenation.WarningCallback = warnings;

        try
        {
            // Register the dictionary from a file. This may throw if the file is missing,
            // unreadable, or has an invalid format.
            Hyphenation.RegisterDictionary(language, dictionaryPath);
            Console.WriteLine($"Dictionary for '{language}' registered successfully.");
        }
        catch (UnsupportedFileFormatException ex)
        {
            // Thrown when the dictionary format is not recognized.
            Console.WriteLine($"Unsupported hyphenation dictionary format: {ex.Message}");
        }
        catch (IOException ex)
        {
            // Thrown for I/O problems such as file not found or access denied.
            Console.WriteLine($"I/O error while reading '{dictionaryPath}': {ex.Message}");
        }
        catch (Exception ex)
        {
            // Catch‑all for any other unexpected errors.
            Console.WriteLine($"Failed to register hyphenation dictionary for '{language}': {ex.Message}");
        }

        // Output any warnings collected during registration.
        if (warnings.Count > 0)
        {
            Console.WriteLine("Warnings during dictionary registration:");
            foreach (WarningInfo info in warnings)
            {
                Console.WriteLine($"- [{info.WarningType}] {info.Description}");
            }
        }
    }

    // Example entry point.
    public static void Main()
    {
        // Example of a dictionary that may be incompatible or corrupted.
        string dictionaryPath = @"C:\Hyphenation\invalid.dic";

        RegisterDictionaryWithHandling("en-US", dictionaryPath);
    }
}
