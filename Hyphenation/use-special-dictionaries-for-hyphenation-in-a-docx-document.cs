using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Paths to input document, output document and hyphenation dictionaries.
        string inputDocPath = @"C:\Docs\German text.docx";
        string outputDocPath = @"C:\Docs\German text hyphenated.docx";
        string enUsDictionaryPath = @"C:\Dictionaries\hyph_en_US.dic";
        string deChDictionaryPath = @"C:\Dictionaries\hyph_de_CH.dic";

        // Register a known dictionary upfront (optional).
        Hyphenation.RegisterDictionary("en-US", enUsDictionaryPath);

        // Set up a callback that will register dictionaries on demand.
        Hyphenation.Callback = new CustomHyphenationDictionaryRegister(new Dictionary<string, string>
        {
            { "en-US", enUsDictionaryPath },
            { "de-CH", deChDictionaryPath }
        });

        // Load the document.
        Document doc = new Document(inputDocPath);

        // Enable automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch.
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenateCaps = true;

        // Save the hyphenated document.
        doc.Save(outputDocPath);
    }
}

// Implements IHyphenationCallback to register dictionaries when requested.
class CustomHyphenationDictionaryRegister : IHyphenationCallback
{
    private readonly Dictionary<string, string> _dictionaryFiles;

    public CustomHyphenationDictionaryRegister(Dictionary<string, string> dictionaryFiles)
    {
        _dictionaryFiles = dictionaryFiles;
    }

    public void RequestDictionary(string language)
    {
        Console.Write($"Hyphenation dictionary requested: {language}");

        if (Hyphenation.IsDictionaryRegistered(language))
        {
            Console.WriteLine(", already registered.");
            return;
        }

        if (_dictionaryFiles.TryGetValue(language, out string filePath))
        {
            Hyphenation.RegisterDictionary(language, filePath);
            Console.WriteLine(", successfully registered.");
        }
        else
        {
            Console.WriteLine(", no dictionary file known for this language.");
        }
    }
}
