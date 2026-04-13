using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Language code for which we want to check hyphenation support.
        const string language = "en-US";

        // Check if a hyphenation dictionary is already registered for the language.
        bool isRegistered = Hyphenation.IsDictionaryRegistered(language);
        Console.WriteLine($"Dictionary registered for {language}: {isRegistered}");

        // If not registered, create a minimal dummy dictionary file and attempt registration.
        // The dummy file is empty; registration may succeed but will not provide real patterns.
        if (!isRegistered)
        {
            try
            {
                string dicPath = Path.Combine(Directory.GetCurrentDirectory(), "hyph_en_US.dic");
                // Create an empty dictionary file (placeholder).
                File.WriteAllText(dicPath, string.Empty);

                using (FileStream stream = new FileStream(dicPath, FileMode.Open, FileAccess.Read))
                {
                    Hyphenation.RegisterDictionary(language, stream);
                }

                isRegistered = Hyphenation.IsDictionaryRegistered(language);
                Console.WriteLine($"After registration attempt, dictionary registered: {isRegistered}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to register dictionary: {ex.Message}");
            }
        }

        // Create a new document and enable automatic hyphenation.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        doc.HyphenationOptions.AutoHyphenation = true;

        // Narrow the page width to force line breaks where hyphenation could occur.
        doc.FirstSection.PageSetup.PageWidth = 200; // points

        // Add text that contains the word "hyphenation".
        builder.Writeln("This is a demonstration of hyphenation. The word hyphenation may be split across lines.");

        // Save the document to the local folder.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "HyphenationDemo.docx");
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");

        // Final result: if a dictionary is registered and AutoHyphenation is true,
        // the word "hyphenation" can be hyphenated during layout.
        if (isRegistered && doc.HyphenationOptions.AutoHyphenation)
        {
            Console.WriteLine("Hyphenation is enabled and a dictionary is available; the word may be hyphenated.");
        }
        else
        {
            Console.WriteLine("Hyphenation is not possible because either the dictionary is missing or AutoHyphenation is disabled.");
        }
    }
}
