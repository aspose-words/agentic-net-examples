using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

class HyphenationExample
{
    static void Main()
    {
        // Paths to the input document and the hyphenation dictionary.
        // Adjust these paths to point to your actual files.
        string dataDir = @"C:\Data\";
        string inputDocPath = Path.Combine(dataDir, "input.docx");
        string dictionaryPath = Path.Combine(dataDir, "hyph_en_US.dic");
        string outputDocPath = Path.Combine(dataDir, "output.docx");

        // Load the existing DOCX file.
        Document doc = new Document(inputDocPath);

        // Register the custom hyphenation dictionary for the desired language (e.g., en-US).
        Hyphenation.RegisterDictionary("en-US", dictionaryPath);

        // Enable automatic hyphenation for the document.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Save the modified document.
        doc.Save(outputDocPath);
    }
}
