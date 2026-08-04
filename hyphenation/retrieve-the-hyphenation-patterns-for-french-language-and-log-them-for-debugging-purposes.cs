using System;
using System.Globalization;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a simple French hyphenation dictionary locally.
        const string dictionaryPath = "hyph_fr_FR.dic";
        const string dictionaryContent =
            "UTF-8\n" +
            "extraordinaire=ex-tra-or-di-nai-re\n" +
            "communication=com-mu-ni-ca-tion\n";

        // Write the dictionary file if it does not already exist.
        if (!File.Exists(dictionaryPath))
            File.WriteAllText(dictionaryPath, dictionaryContent);

        // Register the dictionary with Aspose.Words.
        Aspose.Words.Hyphenation.RegisterDictionary("fr-FR", dictionaryPath);

        // Verify registration succeeded.
        if (!Aspose.Words.Hyphenation.IsDictionaryRegistered("fr-FR"))
            throw new InvalidOperationException("Failed to register the French hyphenation dictionary.");

        // Log the dictionary contents for debugging.
        Console.WriteLine("French hyphenation patterns loaded from dictionary:");
        Console.WriteLine(File.ReadAllText(dictionaryPath));

        // Build a sample document that uses the French locale.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.LocaleId = new CultureInfo("fr-FR").LCID;

        // Write words that have hyphenation points defined in the dictionary.
        builder.Writeln("extraordinaire communication");

        // Enable automatic hyphenation so the patterns are applied during layout.
        doc.HyphenationOptions.AutoHyphenation = true;

        // Save the document to PDF to trigger layout processing.
        const string outputPath = "Hyphenated_fr.pdf";
        doc.Save(outputPath);

        // Ensure the PDF was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The output PDF was not created.");

        Console.WriteLine($"Document saved to '{outputPath}'.");
    }
}
