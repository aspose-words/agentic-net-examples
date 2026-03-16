using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Net.Http;
using System.Threading.Tasks;
using Aspose.Words;

/// <summary>
/// Downloads the latest LibreOffice hyphenation dictionaries, extracts them,
/// and registers them with Aspose.Words so that hyphenation works for any language
/// present in the downloaded set.
/// </summary>
class Program
{
    // URL of the LibreOffice dictionaries repository (zip of the master branch).
    private const string DictionariesZipUrl = "https://github.com/LibreOffice/dictionaries/archive/refs/heads/master.zip";

    // Folder where the dictionaries will be stored after extraction.
    private static readonly string DictionariesFolder = Path.Combine(Path.GetTempPath(), "LibreOfficeHyphenation");

    static async Task Main()
    {
        // Ensure the target folder exists and is clean.
        if (Directory.Exists(DictionariesFolder))
            Directory.Delete(DictionariesFolder, recursive: true);
        Directory.CreateDirectory(DictionariesFolder);

        // Download and extract the dictionaries.
        await DownloadAndExtractDictionariesAsync(DictionariesZipUrl, DictionariesFolder);

        // Set up a callback that will load dictionaries on demand from the extracted files.
        Hyphenation.Callback = new LibreOfficeDictionaryCallback(DictionariesFolder);

        // Example usage: create a document, enable auto‑hyphenation and save it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Size = 24;
        builder.Writeln("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");

        // Turn on automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Save the document – during layout Aspose.Words will request any needed dictionaries.
        string outputPath = Path.Combine(DictionariesFolder, "HyphenatedSample.pdf");
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");
    }

    /// <summary>
    /// Downloads the zip file from the given URL and extracts all *.dic files into the target folder.
    /// </summary>
    private static async Task DownloadAndExtractDictionariesAsync(string zipUrl, string targetFolder)
    {
        using HttpClient client = new HttpClient();
        Console.WriteLine($"Downloading dictionaries from {zipUrl} ...");
        byte[] zipBytes = await client.GetByteArrayAsync(zipUrl);
        Console.WriteLine("Download complete. Extracting...");

        using MemoryStream zipStream = new MemoryStream(zipBytes);
        using ZipArchive archive = new ZipArchive(zipStream, ZipArchiveMode.Read);
        foreach (ZipArchiveEntry entry in archive.Entries)
        {
            // Only process files that end with .dic (LibreOffice dictionary files).
            if (entry.FullName.EndsWith(".dic", StringComparison.OrdinalIgnoreCase))
            {
                // The entry path includes a top‑level folder like "dictionaries-master/". Remove it.
                string fileName = Path.GetFileName(entry.FullName);
                string destinationPath = Path.Combine(targetFolder, fileName);
                entry.ExtractToFile(destinationPath);
                Console.WriteLine($"Extracted: {fileName}");
            }
        }
        Console.WriteLine("Extraction finished.");
    }

    /// <summary>
    /// Implements IHyphenationCallback to load dictionaries from a local folder.
    /// The folder should contain files named like "hyph_en_US.dic".
    /// </summary>
    private class LibreOfficeDictionaryCallback : IHyphenationCallback
    {
        private readonly Dictionary<string, string> _languageToFile = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        public LibreOfficeDictionaryCallback(string dictionariesFolder)
        {
            // Scan the folder for *.dic files and build a map from language code to file path.
            foreach (string filePath in Directory.GetFiles(dictionariesFolder, "*.dic"))
            {
                string fileName = Path.GetFileNameWithoutExtension(filePath); // e.g., hyph_en_US
                // Expected pattern: hyph_<lang>.dic
                if (fileName.StartsWith("hyph_", StringComparison.OrdinalIgnoreCase))
                {
                    string langPart = fileName.Substring(5); // remove "hyph_"
                    // Convert underscores to hyphens to match .NET culture names (en_US -> en-US).
                    string cultureName = langPart.Replace('_', '-');
                    // Some dictionaries use lower‑case language codes; ensure proper case.
                    try
                    {
                        // Validate that the culture exists; if not, we still keep the raw string.
                        CultureInfo.GetCultureInfo(cultureName);
                    }
                    catch (CultureNotFoundException) { /* ignore validation */ }

                    if (!_languageToFile.ContainsKey(cultureName))
                        _languageToFile.Add(cultureName, filePath);
                }
            }
        }

        public void RequestDictionary(string language)
        {
            Console.Write($"Hyphenation dictionary requested: {language}");

            // If already registered, do nothing.
            if (Hyphenation.IsDictionaryRegistered(language))
            {
                Console.WriteLine(", already registered.");
                return;
            }

            // Try to find a matching dictionary file.
            if (_languageToFile.TryGetValue(language, out string filePath))
            {
                Hyphenation.RegisterDictionary(language, filePath);
                Console.WriteLine(", successfully registered.");
                return;
            }

            // No dictionary found – register a null dictionary to avoid repeated callbacks.
            Hyphenation.RegisterDictionary(language, (string)null);
            Console.WriteLine(", not found. Registered null dictionary to suppress further requests.");
        }
    }
}
