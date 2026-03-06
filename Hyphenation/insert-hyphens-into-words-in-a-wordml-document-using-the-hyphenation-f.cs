using System;
using Aspose.Words;

class Program
{
    static void Main()
    {
        // Load the WORDML document.
        string inputPath = "InputDocument.xml";
        Document doc = new Document(inputPath);

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.HyphenationZone = 720; // 0.5 inch from the right margin.
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2; // Max two consecutive hyphenated lines.

        // Register a hyphenation dictionary for the document language if needed.
        // This example uses English (US); adjust the language code and dictionary file as required.
        const string language = "en-US";
        const string dictionaryFile = "hyph_en_US.dic";

        // The Hyphenation class resides directly under the Aspose.Words namespace, not in a separate namespace.
        if (!Aspose.Words.Hyphenation.IsDictionaryRegistered(language))
        {
            Aspose.Words.Hyphenation.RegisterDictionary(language, dictionaryFile);
        }

        // Save the document; hyphens will be applied during layout.
        string outputPath = "OutputDocument.docx";
        doc.Save(outputPath);
    }
}
