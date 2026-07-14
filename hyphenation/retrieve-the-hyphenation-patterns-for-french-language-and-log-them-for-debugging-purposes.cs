using System;
using System.Globalization;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Path for the French hyphenation dictionary.
        const string dictionaryPath = "hyph_fr_FR.dic";

        // Create a minimal French hyphenation dictionary.
        // The first line must be the encoding identifier (e.g., UTF-8).
        // Subsequent lines contain word=hyphenated-pattern entries.
        string dictionaryContent =
            "UTF-8\n" +
            "bonjour=bon-jour\n" +
            "au-revoir=au-re-voir\n" +
            "extraordinairement=ex-tra-or-di-na-i-re-ment";

        File.WriteAllText(dictionaryPath, dictionaryContent);

        // Verify that the dictionary file was created.
        if (!File.Exists(dictionaryPath))
            throw new InvalidOperationException($"Dictionary file '{dictionaryPath}' was not created.");

        // Register the French dictionary with Aspose.Words.
        Hyphenation.RegisterDictionary("fr-FR", dictionaryPath);

        // Confirm that the dictionary is now registered.
        if (!Hyphenation.IsDictionaryRegistered("fr-FR"))
            throw new InvalidOperationException("Failed to register the French hyphenation dictionary.");

        // Retrieve the patterns by reading the dictionary file.
        // In a real scenario you might parse the file; here we simply output its contents.
        string loadedPatterns = File.ReadAllText(dictionaryPath);

        // Log the patterns for debugging purposes.
        Console.WriteLine("French hyphenation patterns loaded from dictionary:");
        Console.WriteLine(loadedPatterns);

        // Demonstrate that the dictionary is usable by creating a document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Size = 24;
        builder.Font.LocaleId = new CultureInfo("fr-FR").LCID;
        builder.Writeln("extraordinairement au-revoir bonjour");

        // Enable automatic hyphenation so the dictionary is applied.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Save the document to verify that hyphenation works (output not required by the task).
        const string outputPath = "HyphenatedFrench.docx";
        doc.Save(outputPath);

        // Validate that the document was saved.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output document was not created.");

        // Optional cleanup (commented out to keep files for inspection).
        // File.Delete(dictionaryPath);
        // File.Delete(outputPath);
    }
}
