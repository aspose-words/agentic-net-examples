using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace HyphenatePdfApp
{
    class Program
    {
        // Entry point: args[0] = input document path,
        //               args[1] = hyphenation language code (e.g. "en-US"),
        //               args[2] = optional output PDF path.
        static void Main(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: HyphenatePdfApp <inputDocPath> <languageCode> [outputPdfPath]");
                return;
            }

            string inputPath = args[0];
            string languageCode = args[1];
            string outputPath = args.Length >= 3 ? args[2] : Path.ChangeExtension(inputPath, ".pdf");

            try
            {
                // Load the source document.
                Document doc = new Document(inputPath);

                // Enable automatic hyphenation.
                doc.HyphenationOptions.AutoHyphenation = true;
                // Optional: configure additional hyphenation settings.
                doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
                doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch
                doc.HyphenationOptions.HyphenateCaps = true;

                // Register the hyphenation dictionary for the requested language if it is not already registered.
                // Expect a dictionary file named "hyph_<languageCode>.dic" in the same folder as the executable.
                if (!Hyphenation.IsDictionaryRegistered(languageCode))
                {
                    string dictFileName = $"hyph_{languageCode}.dic";
                    string dictPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, dictFileName);

                    if (!File.Exists(dictPath))
                    {
                        Console.WriteLine($"Hyphenation dictionary not found: {dictPath}");
                        return;
                    }

                    Hyphenation.RegisterDictionary(languageCode, dictPath);
                }

                // Save the document as PDF.
                doc.Save(outputPath, SaveFormat.Pdf);

                Console.WriteLine($"Hyphenated PDF saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
