using System;
using System.Globalization;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;

public class Program
{
    public static void Main()
    {
        // Prepare a folder for temporary files.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "HyphenationDemo");
        Directory.CreateDirectory(workDir);

        // Create minimal hyphenation dictionary files for English (US) and German (Switzerland).
        // The OpenOffice hyphenation dictionary format requires at least a header line.
        // For demonstration purposes an empty pattern list is sufficient to register the dictionary.
        string enDictPath = Path.Combine(workDir, "hyph_en_US.dic");
        string deDictPath = Path.Combine(workDir, "hyph_de_CH.dic");
        File.WriteAllText(enDictPath, "SET UTF-8\n");
        File.WriteAllText(deDictPath, "SET UTF-8\n");

        // Register the dictionaries so that Aspose.Words can hyphenate the corresponding languages.
        using (FileStream enStream = File.OpenRead(enDictPath))
        {
            Hyphenation.RegisterDictionary("en-US", enStream);
        }
        using (FileStream deStream = File.OpenRead(deDictPath))
        {
            Hyphenation.RegisterDictionary("de-CH", deStream);
        }

        // Verify that the dictionaries are registered.
        if (!Hyphenation.IsDictionaryRegistered("en-US") || !Hyphenation.IsDictionaryRegistered("de-CH"))
            throw new InvalidOperationException("Failed to register hyphenation dictionaries.");

        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Narrow the page width to force line wrapping and enable hyphenation.
        doc.FirstSection.PageSetup.PageWidth = 300; // points (~4.2 cm)
        doc.FirstSection.PageSetup.PageHeight = 842; // A4 height for completeness.

        // Enable automatic hyphenation for the whole document.
        doc.HyphenationOptions.AutoHyphenation = true;
        doc.HyphenationOptions.HyphenateCaps = true;
        doc.HyphenationOptions.ConsecutiveHyphenLimit = 2;
        doc.HyphenationOptions.HyphenationZone = 360; // default

        // English paragraph (en-US).
        builder.Font.Size = 12;
        builder.Font.LocaleId = new CultureInfo("en-US").LCID;
        builder.Writeln(
            "Hyphenation demonstration with a long English paragraph that should wrap and hyphenate words like demonstration and hyphenation to illustrate the effect.");

        // German paragraph (de-CH).
        builder.Font.Size = 12;
        builder.Font.LocaleId = new CultureInfo("de-CH").LCID;
        builder.Writeln(
            "Demonstration der Silbentrennung mit einem langen deutschen Absatz, der Zeilen umbrechen und Wörter wie Demonstration und Silbentrennung trennen sollte, um den Effekt zu zeigen.");

        // Save the document to PDF – PDF rendering forces layout and applies hyphenation.
        string outPath = Path.Combine(workDir, "HyphenationMixedLanguages.pdf");
        doc.Save(outPath, SaveFormat.Pdf);

        // Validate that the output file was created.
        if (!File.Exists(outPath))
            throw new FileNotFoundException("The PDF output was not created.", outPath);
    }
}
