using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

class HyphenationExample
{
    static void Main()
    {
        // Create a temporary working directory.
        string tempDir = Path.Combine(Path.GetTempPath(), "HyphenationExample");
        Directory.CreateDirectory(tempDir);

        // Define file paths.
        string inputDocPath = Path.Combine(tempDir, "input.docx");
        string dictionaryPath = Path.Combine(tempDir, "hyph_en_US.dic");
        string outputDocPath = Path.Combine(tempDir, "output.docx");

        // Create a minimal Hunspell dictionary file.
        // The first line is the number of entries, followed by the words.
        File.WriteAllText(dictionaryPath, "2\nhyphenation\nexample");

        // Create a simple DOCX document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is an example text to demonstrate automatic hyphenation using a custom dictionary.");
        doc.Save(inputDocPath);

        // Load the document (could also continue using the same instance).
        doc = new Document(inputDocPath);

        // Register the custom hyphenation dictionary for en-US.
        Hyphenation.RegisterDictionary("en-US", dictionaryPath);

        // Enable automatic hyphenation.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Save the modified document.
        doc.Save(outputDocPath);

        Console.WriteLine($"Input document: {inputDocPath}");
        Console.WriteLine($"Dictionary file: {dictionaryPath}");
        Console.WriteLine($"Output document: {outputDocPath}");
    }
}
